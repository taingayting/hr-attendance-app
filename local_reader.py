"""Free photo reader: no API, no limits. Straighten the photo, read the printed IDs with OCR, find the pen ticks by colour,
and say which cell each tick is in. Anything unclear is marked, never guessed."""
import re, math
import numpy as np, cv2

OT_ORDER = ["0.5", "1", "2", "2.5", "3", "3.5", "4"]
_ID_RE = re.compile(r"^[XK]?HC[-—–_.:]?([0-9OolIBSZ]{5})(?![0-9])", re.I)   # text after the ID (the Khmer name read as junk) is allowed
_DIG = str.maketrans({"O": "0", "o": "0", "l": "1", "I": "1", "B": "8", "S": "5", "Z": "2"})
_OCR = None


def get_ocr():
    global _OCR
    if _OCR is None:
        from rapidocr_onnxruntime import RapidOCR
        _OCR = RapidOCR()
    return _OCR


def _norm_w(im, w=1500):
    k = w / im.shape[1]
    return cv2.resize(im, None, fx=k, fy=k, interpolation=cv2.INTER_AREA if k < 1 else cv2.INTER_CUBIC)


def _deskew(im):
    g = cv2.cvtColor(im, cv2.COLOR_BGR2GRAY)
    bw = cv2.adaptiveThreshold(g, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 25, 12)
    lines = cv2.HoughLinesP(bw, 1, np.pi / 720, threshold=200, minLineLength=im.shape[1] // 4, maxLineGap=15)
    ang = []
    if lines is not None:
        for x1, y1, x2, y2 in lines.reshape(-1, 4):
            a = math.degrees(math.atan2(y2 - y1, x2 - x1))
            if abs(a) < 12:
                ang.append(a)
    a = float(np.median(ang)) if ang else 0.0
    h, w = im.shape[:2]
    M = cv2.getRotationMatrix2D((w / 2, h / 2), a, 1.0)
    return cv2.warpAffine(im, M, (w, h), flags=cv2.INTER_CUBIC, borderValue=(255, 255, 255)), a


def _ocr_items(ocr, im):
    res, _ = ocr(im)
    out = []
    for box, txt, conf in (res or []):
        xs = [p[0] for p in box]; ys = [p[1] for p in box]
        out.append(dict(cx=sum(xs) / 4, cy=sum(ys) / 4, w=max(xs) - min(xs), h=max(ys) - min(ys), t=str(txt), c=float(conf)))
    return out


_MULTI = re.compile(r"(0\.5|2\.5|3\.5|[1-4])H")


def _split_labels(it):
    """OCR often reads the whole row '0.5H 1H 2H 2.5H 3H 3.5H' as ONE text. Cut it into the separate headings and
    place each one along the box by its character position."""
    s = re.sub(r"[\s。,，]", "", it["t"]).upper().replace("O", "0").replace("．", ".")
    ms = list(_MULTI.finditer(s))
    if len(ms) < 2 or "".join(m.group(0) for m in ms) != s:
        return None
    out, n = [], len(s)
    x0 = it["cx"] - it["w"] / 2
    for m in ms:
        mid = (m.start() + m.end()) / 2 / n
        out.append((dict(it, cx=x0 + it["w"] * mid, w=it["w"] * (m.end() - m.start()) / n, t=m.group(0)), m.group(1)))
    return out


def _label(t):
    s = re.sub(r"[\s。,，]", "", t).upper().replace("O", "0").replace("I", "1").replace("L", "1").replace("．", ".")
    m = re.fullmatch(r"(0\.5|1|2|2\.5|3|3\.5|4|8)H|H(0\.5|1|2|2\.5|3|3\.5|4|8)", s)
    return (m.group(1) or m.group(2)) if m else None


def _blue(im):
    """Pen ink = dark strokes that are NOT part of the printed grid (works for blue, black and purple pens)."""
    g = cv2.cvtColor(im, cv2.COLOR_BGR2GRAY)
    g = cv2.GaussianBlur(g, (3, 3), 0)
    dark = cv2.adaptiveThreshold(g, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 41, 14)
    hk = cv2.morphologyEx(dark, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (50, 1)))
    vk = cv2.morphologyEx(dark, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, 40)))
    grid = cv2.dilate(cv2.bitwise_or(hk, vk), cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5)))
    ink = cv2.bitwise_and(dark, cv2.bitwise_not(grid))
    ink = cv2.morphologyEx(ink, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2)))
    return ink


def _interp_cols(found):
    """found: {label: x}. Fill the missing OT labels from their neighbours."""
    xs = {}
    idx = [i for i, l in enumerate(OT_ORDER) if l in found]
    if len(idx) < 3:
        return None
    for i, l in enumerate(OT_ORDER):
        if l in found:
            xs[l] = found[l]
        else:
            lo = max([j for j in idx if j < i], default=None); hi = min([j for j in idx if j > i], default=None)
            if lo is not None and hi is not None:
                xs[l] = found[OT_ORDER[lo]] + (found[OT_ORDER[hi]] - found[OT_ORDER[lo]]) * (i - lo) / (hi - lo)
            elif lo is not None:
                step = (found[OT_ORDER[idx[-1]]] - found[OT_ORDER[idx[0]]]) / max(1, idx[-1] - idx[0])
                xs[l] = found[OT_ORDER[lo]] + step * (i - lo)
            else:
                step = (found[OT_ORDER[idx[-1]]] - found[OT_ORDER[idx[0]]]) / max(1, idx[-1] - idx[0])
                xs[l] = found[OT_ORDER[hi]] - step * (hi - i)
    return xs


def vertical_tilt(im):
    """dx/dy of the vertical table lines (perspective makes them lean a little)."""
    g = cv2.cvtColor(im, cv2.COLOR_BGR2GRAY)
    bw = cv2.adaptiveThreshold(g, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 25, 12)
    lines = cv2.HoughLinesP(bw, 1, np.pi / 720, threshold=120, minLineLength=im.shape[0] // 8, maxLineGap=10)
    t = []
    if lines is not None:
        for x1, y1, x2, y2 in lines.reshape(-1, 4):
            if abs(y2 - y1) > abs(x2 - x1) * 4:
                t.append((x2 - x1) / (y2 - y1))
    return float(np.median(t)) if len(t) >= 3 else 0.0


def _score(res):
    """How believable a reading is: a 'Came' mark with NO 8H tick is the typical sign that a tick was missed (a bad picture size)."""
    rows = res.get("rows", [])
    good = sum(1 for r in rows if r["came"] and r["h8"])
    odd = sum(1 for r in rows if (r["came"] and not r["h8"] and not r["ot"]) or (r["h8"] and not r["came"]))
    ids_ = [r["id"] for r in rows if r["id"]]
    dups = len(ids_) - len(set(ids_))                                                   # the same ID twice = a misread ID
    noid = sum(1 for r in rows if not r["id"] and (r["came"] or r["h8"] or r["ot"]))    # ticks but no ID
    return good - 2 * odd - 2 * dups - noid


def read_sheet(bgr, roster=None, ocr=None, debug=None):
    """Try a few picture sizes: the heading OCR and the tick finder are sensitive to small changes in the picture.
    A reading with nothing odd AND ticks on at least half of the rows wins at once; otherwise every size is tried and the most believable wins."""
    last, best = None, None
    cropped = _crop_paper(bgr)
    same = cropped.shape == bgr.shape
    for crop in ((True,) if same else (True, False)):
        src = cropped if crop else bgr
        for w in (1500, 1800, 1250, 2100):
            last = _read_sheet(src, roster, ocr, debug, w)
            if last["ok"]:
                sc = _score(last)
                good = sum(1 for r in last["rows"] if r["came"] and r["h8"])
                ids_ = [r["id"] for r in last["rows"] if r["id"]]
                noid_ = sum(1 for r in last["rows"] if not r["id"] and (r["came"] or r["h8"] or r["ot"]))     # a handwritten new worker has no printed ID: not a reason to try again
                if sc + noid_ == good and len(ids_) == len(set(ids_)) and last["rows"] and good >= 0.5 * len(last["rows"]):      # nothing odd, no ID twice, most rows have ticks
                    return last
                if best is None or (sc, good) > (best[0], best[2]):
                    best = (sc, last, good)
            elif "heading" not in last.get("error", ""):
                return last
    return best[1] if best else last


def _crop_paper(bgr):
    """Cut away the table/desk around the paper so the table fills the picture (the reader works much better on a close picture)."""
    g = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    small = cv2.resize(g, None, fx=0.25, fy=0.25)
    blur = cv2.GaussianBlur(small, (9, 9), 0)
    _, bw = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    bw = cv2.morphologyEx(bw, cv2.MORPH_CLOSE, cv2.getStructuringElement(cv2.MORPH_RECT, (15, 15)))
    cs, _ = cv2.findContours(bw, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not cs:
        return bgr
    c = max(cs, key=cv2.contourArea)
    x, y, w, h = cv2.boundingRect(c)
    H, W = small.shape
    if cv2.contourArea(c) < 0.15 * W * H or (w * h) > 0.85 * W * H:
        return bgr
    m = 6
    x0, y0, x1, y1 = max(0, x - m), max(0, y - m), min(W, x + w + m), min(H, y + h + m)
    return bgr[y0 * 4:y1 * 4, x0 * 4:x1 * 4]


_HEAD_WORDS = ("编号", "工号", "名字", "上班", "备注", "加班")


def _labels_from(items):
    labs = []
    for it in items:
        sp = _split_labels(it)
        if sp:
            labs.extend(sp)
        else:
            l_ = _label(it["t"])
            if l_:
                labs.append((it, l_))
    return labs


def _head_y(items):
    """The height of the printed heading row, found from its Chinese words (编号 工号 名字 是否上班 备注)."""
    ys = [it["cy"] for it in items if any(w in it["t"] for w in _HEAD_WORDS) and "登记" not in it["t"] and len(it["t"]) <= 6]      # not the title 上班登记表
    return float(np.median(ys)) if ys else None


def _settle(labs):
    """From the heading texts: (heading height, the labels on that line, {label: (x, confidence)}, columns or None)."""
    if not labs:
        return None, [], {}, None
    y8 = float(np.median([it["cy"] for it, l in labs if l == "8"])) if any(l == "8" for _, l in labs) else float(np.median([it["cy"] for it, l in labs]))
    labs = [(it, l) for it, l in labs if abs(it["cy"] - y8) < 25]
    found = {}
    for it, l in labs:
        if l != "8" and (l not in found or it["c"] > found[l][1]):
            found[l] = (it["cx"], it["c"])
    return y8, labs, found, _interp_cols({l: v[0] for l, v in found.items()})


def _vline_xs(im):
    """x of the long vertical lines of the table."""
    g = cv2.GaussianBlur(cv2.cvtColor(im, cv2.COLOR_BGR2GRAY), (3, 3), 0)
    dark = cv2.adaptiveThreshold(g, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 41, 14)
    vk = cv2.morphologyEx(dark, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(60, im.shape[0] // 8))))
    prof = (vk > 0).sum(axis=0).astype(float)
    if prof.max() <= 0:
        return []
    thr = max(0.3 * prof.max(), 40)
    xs, i = [], 0
    while i < len(prof):
        if prof[i] >= thr:
            j = i
            while j + 1 < len(prof) and prof[j + 1] >= thr:
                j += 1
            xs.append((i + j) / 2.0)
            i = j + 1
        else:
            i += 1
    out = []
    for x in xs:
        if out and x - out[-1] < 10:
            out[-1] = (out[-1] + x) / 2
        else:
            out.append(x)
    return out


def _grid_cols(im, labs, y8):
    """Safety net: the headings could not be read, so the columns are placed from the vertical lines of the table.
    The 8 cells 8H, 0.5H ... 4H sit side by side; a heading that WAS read anchors them, otherwise they are the 8 cells just left of the wide Remark cell.
    Returns ({label: x}, x of the 8H column) or None."""
    xs = _vline_xs(im)
    if len(xs) < 10:
        return None
    cells = [(xs[i], xs[i + 1]) for i in range(len(xs) - 1)]
    order = ["8"] + OT_ORDER
    bases = set()
    for it, l in labs:
        if l in order:
            ci = next((i for i, (a_, b_) in enumerate(cells) if a_ < it["cx"] < b_), None)
            if ci is not None:
                bases.add(ci - order.index(l))
    if len(bases) > 1:
        return None
    if bases:
        base = bases.pop()
    else:
        wid = [b_ - a_ for a_, b_ in cells]
        med = float(np.median(wid))
        wide = [i for i, w_ in enumerate(wid) if w_ >= 2.0 * med]
        if not wide:
            return None
        base = max(wide) - 8
    if base < 0 or base + 8 > len(cells):
        return None
    mine = cells[base:base + 8]
    wid8 = [b_ - a_ for a_, b_ in mine]
    med8 = float(np.median(wid8))
    if med8 < 25 or any(w_ < 0.55 * med8 or w_ > 1.6 * med8 for w_ in wid8):
        return None
    ctr = [(a_ + b_) / 2.0 for a_, b_ in mine]
    return {l: ctr[i + 1] for i, l in enumerate(OT_ORDER)}, ctr[0]


def _read_sheet(bgr, roster=None, ocr=None, debug=None, width=1500):
    """Returns dict(rows=[{no,id,came,h8,ot,doubt,...}], info, ok, error)."""
    ocr = ocr or get_ocr()
    im = _norm_w(bgr, width)
    im, _ang = _deskew(im)
    items = _ocr_items(ocr, im)
    tilt = vertical_tilt(im)
    # ---- header labels
    labs = _labels_from(items)
    hy = _head_y(items)
    y8, labs, found, cols = _settle(labs)
    h8 = [it for it, l in labs if l == "8"]
    h8_guess = False
    grid_guess = False
    if cols is None and (y8 is not None or hy is not None):
        yy = y8 if y8 is not None else hy
        # safety net 1: the heading strip is read again, enlarged (small or faint print)
        y0b = int(max(0, yy - 45)); y1b = int(min(im.shape[0], yy + 45))
        if y1b - y0b > 20:
            big = cv2.resize(im[y0b:y1b, :], None, fx=2.5, fy=2.5, interpolation=cv2.INTER_CUBIC)
            try:
                its2 = _ocr_items(ocr, big)
            except Exception:
                its2 = []
            for it2 in its2:
                it2.update(cx=it2["cx"] / 2.5, cy=it2["cy"] / 2.5 + y0b, w=it2["w"] / 2.5, h=it2["h"] / 2.5)
            y8b, labs_b, found_b, cols_b = _settle(_labels_from(items) + _labels_from(its2))
            if cols_b is not None:
                items = items + its2
                y8, labs, found, cols = y8b, labs_b, found_b, cols_b
                h8 = [it for it, l in labs if l == "8"]
        # safety net 2: the columns are placed from the vertical lines of the table
        if cols is None:
            g_ = _grid_cols(im, labs, yy)
            if g_:
                cols, x8g = g_
                y8 = yy
                h8 = [dict(cx=float(x8g), cy=float(yy), w=40.0, h=20.0, t="8H", c=0.0)]
                h8_guess = grid_guess = True
    if y8 is None and cols is None:
        return dict(ok=False, error="could not find the column headings (0.5H, 1H, ...). Is the whole table in the photo?", rows=[], info={})
    if cols is not None and not h8:
        # the OCR often drops the '8H' heading (it is printed with a line under it). Place the 8H column from the other headings:
        # it sits between the 'Came' column and the 0.5H column.
        cl_ = [it for it in items if "上班" in it["t"] and abs(it["cy"] - y8) < 30]
        if cl_:
            x8e = cl_[0]["cx"] + 0.56 * (cols["0.5"] - cl_[0]["cx"])
        else:
            x8e = cols["0.5"] - 1.33 * (cols["4"] - cols["0.5"]) / 6.0
        h8 = [dict(cx=float(x8e), cy=float(y8), w=40.0, h=20.0, t="8H", c=0.0)]
        h8_guess = True
    if cols is None or not h8:
        return dict(ok=False, error="could not find the 8H / OT column headings", rows=[], info={})
    x8 = h8[0]["cx"]
    came_lab = [it for it in items if "上班" in it["t"] and abs(it["cy"] - y8) < 30]
    xcame = came_lab[0]["cx"] if came_lab else x8 - 0.95 * (cols["0.5"] - x8) * 2.0
    head_y = float(y8)
    # ---- printed IDs
    ids = []
    for it in items:
        s = re.sub(r"[\s]", "", it["t"])
        m = _ID_RE.match(s)
        if m and it["cy"] > head_y + 10:
            cxi = it["cx"] if m.end() >= len(s) else it["cx"] - it["w"] / 2 + 62   # the ID is at the left of a long text
            ids.append(dict(it, cx=cxi, id="DMO-" + m.group(1).translate(_DIG)))
    if not ids:
        return dict(ok=False, error="no printed worker IDs found", rows=[], info={})
    xid = float(np.median([i["cx"] for i in ids]))
    ids.sort(key=lambda i: i["cy"])
    diffs = [b["cy"] - a["cy"] for a, b in zip(ids, ids[1:]) if 0 < b["cy"] - a["cy"] < 90]
    pitch = float(np.median(diffs)) if diffs else 30.0
    # row index k for each ID by rounding on a growing grid
    y_first = ids[0]["cy"]
    for i in ids:
        i["k"] = int(round((i["cy"] - y_first) / pitch))
    # refine pitch by linear fit
    ks = np.array([i["k"] for i in ids], float); ys = np.array([i["cy"] for i in ids], float)
    if len(set(ks)) >= 3:
        pitch, y0 = np.polyfit(ks, ys, 1)
    else:
        y0 = y_first
    nrows = int(max(i["k"] for i in ids)) + 1
    # ---- an ID cell the first OCR pass skipped: read just that cell again, enlarged
    have_k = {i["k"] for i in ids}
    for k_ in range(nrows):
        if k_ in have_k:
            continue
        yc_ = y0 + pitch * k_
        c0_, c1_ = int(max(0, xid - 95)), int(min(im.shape[1], xid + 95))
        for pad_ in (0.75, 0.55):                             # a taller cell first: the 'DMO-' and the digits then stay in one piece
            r0_, r1_ = int(max(0, yc_ - pitch * pad_)), int(min(im.shape[0], yc_ + pitch * pad_))
            if r1_ - r0_ < 10 or c1_ - c0_ < 10:
                continue
            strip = cv2.resize(im[r0_:r1_, c0_:c1_], None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC)
            try:
                mid_ = (yc_ - r0_) * 3                          # the middle of THIS row inside the enlarged strip
                its_ = sorted([t for t in _ocr_items(ocr, strip) if abs(t["cy"] - mid_) < 0.4 * pitch * 3], key=lambda t: t["cx"])   # text of the rows above / below is ignored
            except Exception:
                continue
            hit_ = None
            for n_, t_ in enumerate(its_):
                s_ = re.sub(r"[\s]", "", t_["t"])
                m_ = _ID_RE.match(s_)
                if m_ and t_["c"] >= 0.5:
                    hit_ = (m_.group(1), t_["c"]); break
                if re.fullmatch(r"[0-9OolIBSZ]{5}", s_) and n_ > 0 and re.match(r"^[XKH]", re.sub(r"[\s]", "", its_[n_ - 1]["t"]).upper()) and t_["c"] >= 0.5:
                    hit_ = (s_, min(t_["c"], its_[n_ - 1]["c"])); break          # 'XHIC-1' + '10093' read as two pieces
            if hit_ and any(i_["id"] == "DMO-" + hit_[0].translate(_DIG) for i_ in ids):
                hit_ = None                                     # that ID is already on another row: it is the neighbour's, not this row's
            if hit_:
                ids.append(dict(cx=xid, cy=float(yc_), c=hit_[1], k=k_, id="DMO-" + hit_[0].translate(_DIG), t=hit_[0], w=60.0, h=20.0))
                have_k.add(k_)
                break
    # ---- ink
    mask = _blue(im)
    xleft = xid + 0.9 * (xcame - xid) * 0  # not used
    x_ink_right = cols["4"] + 0.6 * (cols["4"] - cols["3.5"])
    # extend rows below the last printed ID while the ID cell has handwriting
    id_x0, id_x1 = int(xid - 75), int(xid + 75)
    def row_y(k, x=None):
        return y0 + pitch * k
    k = nrows
    extra = 0
    while k < nrows + 12:
        yc = row_y(k); r0, r1 = int(yc - pitch * .4), int(yc + pitch * .4)
        if r1 >= im.shape[0]:
            break
        ink = (mask[r0:r1, id_x0:id_x1] > 0).sum()
        if ink > 60:
            extra = k - nrows + 1
        elif k - nrows - extra >= 1:
            break
        k += 1
    nrows += extra
    # ---- tick blobs
    mm = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (9, 9)))
    n, lab, stats, cent = cv2.connectedComponentsWithStats(mm, 8)
    raw_n, raw_lab, raw_stats, _ = cv2.connectedComponentsWithStats(mask, 8)
    blobs = []
    xlo = xcame - 60
    for i in range(1, n):
        x, y, w, h, a = stats[i]
        cx, cy = cent[i]
        if a < 40 or cx < xlo or cx > x_ink_right + 30:
            continue
        if h > 1.3 * pitch and w < 2.2 * pitch:                  # a big check mark: its long stroke reaches into the row above. The row is where it ENDS.
            ncomp_ = len({v for v in np.unique(raw_lab[y:y + h, x:x + w][(lab[y:y + h, x:x + w] == i)]) if v != 0 and raw_stats[v][4] > 8})
            if h > 1.6 * pitch and cx > (x8 + cols["0.5"]) / 2 and ncomp_ <= 5:
                # in an OT column a tick is about one row tall: a blob nearly two rows tall is TWO ticks (one under the other) joined together
                for dk_ in (0, 1):
                    blobs.append(dict(cx=float(cx), cy=float(y + h - 0.45 * pitch - dk_ * pitch), w=int(w), h=int(pitch), area=int(a / 2), ncomp=2, fill=0.1))
                continue
            cy = y + h - 0.45 * pitch
        if h < 9 or (w > 3.5 * h and h < 14) or (h > 3.5 * w and w < 14):
            continue  # leftover piece of a grid line, not a tick
        if cy < head_y + 12:
            continue
        sub = (lab[y:y + h, x:x + w] == i)
        ncomp = len({v for v in np.unique(raw_lab[y:y + h, x:x + w][sub]) if v != 0 and raw_stats[v][4] > 8})
        fill = float((mask[y:y + h, x:x + w][sub] > 0).sum()) / max(1, w * h)
        blobs.append(dict(cx=float(cx), cy=float(cy), w=int(w), h=int(h), area=int(a), ncomp=ncomp, fill=fill))
    # ---- assign blobs to rows using Came/8H ticks to learn the local slope
    def nearest_k(cy, dy=0.0):
        return int(round((cy - dy - y0) / pitch))
    xs_ref = [b for b in blobs if abs(b["cx"] - xcame) < 45 or abs(b["cx"] - x8) < 40]
    offs = []
    for b in xs_ref:
        kk = nearest_k(b["cy"])
        if 0 <= kk < nrows:
            offs.append((b["cx"], b["cy"] - (y0 + pitch * kk)))
    a_off, b_off = 0.0, 0.0
    if len(offs) >= 6:
        X = np.array([o[0] for o in offs]); Y = np.array([o[1] for o in offs])
        for _ in range(3):
            A = np.vstack([np.ones_like(X), X - xcame]).T
            coef, *_ = np.linalg.lstsq(A, Y, rcond=None)
            res_ = Y - A @ coef
            keep = np.abs(res_) < max(8, 1.5 * np.std(res_))
            if keep.sum() >= 5:
                coef, *_ = np.linalg.lstsq(A[keep], Y[keep], rcond=None)
            a_off, b_off = coef
    def k_of(b):
        dy = a_off + b_off * (b["cx"] - xcame)
        return int(round((b["cy"] - dy - y0) / pitch)), (b["cy"] - dy - y0) / pitch
    colw = {l: None for l in OT_ORDER}
    ot_x = [cols[l] for l in OT_ORDER]
    def col_of(b, k):
        # undo the lean of the vertical lines: x at header height
        x = b["cx"] - tilt * (b["cy"] - head_y)
        cands = [("came", xcame), ("8H", x8)] + [(l, cols[l]) for l in OT_ORDER]
        j = min(range(len(cands)), key=lambda i: abs(cands[i][1] - x))
        # margin to the next column = how sure we are
        dists = sorted(abs(c[1] - x) for c in cands)
        return cands[j][0], (dists[1] - dists[0]) if len(dists) > 1 else 99
    cells = {}
    for b in blobs:
        kk, frac = k_of(b)
        if not (0 <= kk < nrows):
            continue
        cname, margin = col_of(b, kk)
        cells.setdefault(kk, []).append((cname, b, abs(frac - round(frac)), margin))
    # ---- build rows
    id_by_k = {i["k"]: i for i in ids}
    rows = []
    notes = []
    for kk in range(nrows):
        it = id_by_k.get(kk)
        rec = dict(no=kk + 1, id=it["id"] if it else "", id_conf=it["c"] if it else 0.0, came=False, h8=False, ot=[], doubt=[], cy=y0 + pitch * kk)
        cl = cells.get(kk, [])
        # handwriting / many strokes in OT cells
        for cn_ in ("came", "8H"):                      # handwriting (HP, HH, 4.5 ...) in the Came / 8H box is not a tick: several pieces, or a very big blob
            bl_ = [c_[1] for c_ in cl if c_[0] == cn_]
            if len(bl_) >= 2 or any(x_["area"] > 1000 for x_ in bl_):
                rec["hand"] = True
        if rec.get("hand"):
            rec["doubt"].append("handwriting in the Came / 8H box (HP? HH? hours?): read it on the paper")
        for cname, b, off, margin in cl:
            if rec.get("hand") and cname in ("came", "8H"):
                if cname == "came":
                    rec["came"] = True
                continue
            if cname == "came":
                rec["came"] = True
            elif cname == "8H":
                rec["h8"] = True
            else:
                if b["w"] < 18 and b["area"] < 100:
                    continue        # a speck of dust or a bit of grid line, far smaller than a real tick
                wide = b["w"] > 2.6 * max(20, (cols["4"] - cols["0.5"]) / 6)
                if wide or b["ncomp"] >= 4:
                    rec["doubt"].append("handwriting in the OT area (note or number): read it on the paper")
                    rec["ot"].append("?")
                    continue
                if b["ncomp"] == 3 and b["h"] > 16:
                    rec["doubt"].append("handwritten number in the OT area: read it on the paper")
                    rec["ot"].append("?")
                    continue
                rec["ot"].append(cname)
                if b["fill"] > 0.55 and b["area"] > 70:
                    rec["doubt"].append(f"solid ink dot in the {cname}H box (a mark, not a tick?)")
                if margin < 8:
                    rec["doubt"].append(f"tick is close to the line of the next column ({cname}H)")
                if off > 0.38:
                    rec["doubt"].append("a tick is near the line between two rows")
        if not it:
            rec["doubt"].append("ID not printed here (handwritten?): type it")
        rows.append(rec)
    header = [it for it in items if it["cy"] < head_y + 5]
    return dict(ok=True, rows=rows, info=dict(pitch=float(pitch), nrows=nrows, tilt=tilt, h8_guess=h8_guess, grid_guess=grid_guess), cols=cols, x8=x8, xcame=xcame, im=im, header=header, dbg=dict(blobs=blobs, ids=[(i['id'], round(i['cx']), round(i['cy']), i['k']) for i in ids], y0=float(y0), pitch=float(pitch), cells={k: [(c[0], round(c[1]['cx']), round(c[1]['cy'])) for c in v] for k, v in cells.items()}))


def to_table(res):
    """Rows -> (list of dict ID/HOURS/OT/Doubt, not_counted)."""
    out, skipped = [], []
    for r in res["rows"]:
        if not r["id"] and not (r["came"] or r["h8"] or r["ot"]):
            continue
        ot = sorted({x for x in r["ot"] if x != "?"}, key=lambda s: float(s))
        ots = "/".join(ot) + ("/?" if "?" in r["ot"] else "")
        if r["ot"] == ["?"]:
            ots = "?"
        doubt = list(r["doubt"])
        if r.get("hand"):
            out.append(dict(No=r["no"], ID=r["id"], HOURS="?", OT="", Doubt="; ".join(dict.fromkeys(doubt)))); continue
        if not r["came"] and not r["h8"] and not r["ot"]:
            skipped.append((r["no"], r["id"], "no ticks (absent, crossed out or whited out)")); continue
        if r["came"] and not r["h8"] and not r["ot"]:
            skipped.append((r["no"], r["id"], "Came mark but nothing else (letter P = leave?)")); continue
        if len({x for x in r["ot"] if x != "?"}) > 1:
            doubt.append("more than one OT tick found: check which one is real")
        if not r["came"] and (r["h8"] or r["ot"]):
            doubt.append("ticks but no mark in the Came column")
        if r["came"] and not r["h8"] and r["ot"]:
            doubt.append("8H box is empty: is it a full 8 hours?")
        out.append(dict(No=r["no"], ID=r["id"], HOURS="8" if r["h8"] else "", OT=ots, Doubt="; ".join(dict.fromkeys(doubt))))
    return out, skipped


# ------------------------------------------------------------------ adapter for the website
def available():
    """True if opencv and rapidocr are installed (the free reader can run)."""
    try:
        import cv2 as _c, rapidocr_onnxruntime as _r  # noqa: F401
        return True
    except Exception:
        return False


def _variants(bgr):
    """Other versions of the same photo, from the most likely to the least likely to help."""
    yield "turned a quarter to the right", cv2.rotate(bgr, cv2.ROTATE_90_CLOCKWISE)
    yield "turned a quarter to the left", cv2.rotate(bgr, cv2.ROTATE_90_COUNTERCLOCKWISE)
    yield "upside down", cv2.rotate(bgr, cv2.ROTATE_180)
    lab = cv2.cvtColor(bgr, cv2.COLOR_BGR2LAB)
    l_, a_, b_ = cv2.split(lab)
    l2 = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8)).apply(l_)
    yield "with more contrast", cv2.cvtColor(cv2.merge((l2, a_, b_)), cv2.COLOR_LAB2BGR)
    big = cv2.resize(bgr, None, fx=1.5, fy=1.5, interpolation=cv2.INTER_CUBIC)
    blur = cv2.GaussianBlur(big, (0, 0), 3)
    yield "sharper", cv2.addWeighted(big, 1.6, blur, -0.6, 0)


def _rescue(bgr):
    best = None
    for _name, v in _variants(bgr):
        try:
            r = read_sheet(v)
            if r["ok"] and r["rows"]:
                r.setdefault("info", {})["rescued"] = _name
                n_id = sum(1 for x in r["rows"] if x["id"])
                if best is None or n_id > best[0]:
                    best = (n_id, r)
                if n_id >= 8:
                    return r                              # plenty of IDs read: this is the right way up
        except Exception:
            continue
    return best[1] if best and best[0] >= 3 else None


def read_photo_local(jpeg_bytes, fp_ids=None):
    """Same answer shape as paper_reader.read_photo, so the website treats both readers the same. Free, offline, no limit."""
    arr = np.frombuffer(jpeg_bytes, np.uint8)
    bgr = cv2.imdecode(arr, cv2.IMREAD_COLOR)
    if bgr is None:
        raise ValueError("The picture could not be opened.")
    res = read_sheet(bgr)
    few = res["ok"] and len(to_table(res)[0]) < 3
    if few or (not res["ok"] and "heading" in res.get("error", "")):
        res = _rescue(bgr) or res                         # a clear photo that still fails (or gives almost nothing): try it turned, sharper and with more contrast
    if not res["ok"]:
        raise ValueError(res.get("error", "could not read this photo"))
    tab, skipped = to_table(res)
    # the date, team and day/night written at the top of the sheet (so the app can warn about the wrong day)
    htxt = [it["t"] for it in sorted(res.get("header", []), key=lambda i: (round(i["cy"] / 25), i["cx"]))]
    date_text, team, night = "", "", "?"
    for t in htxt:
        m = re.search(r"20\d\d\s*年.{0,4}?\d{1,2}\s*月.{0,4}?\d{1,2}", t) or re.search(r"\d{1,2}\s*[/.\-]\s*\d{1,2}\s*[/.\-]\s*20\d\d", t)
        if m and not date_text:
            date_text = re.sub(r"\s+", "", m.group(0))
        if "夜班" in t:
            night = "yes"
        elif "白班" in t and night == "?":
            night = "no"
        if "/" in t and any(w in t for w in ("切", "印", "仓", "库", "胶", "糊", "品", "检", "装", "AMT", "STOCK", "QC", "Cut", "Print", "Glue")) and not team:
            team = t.strip()[:40]
    rows = []
    for r in tab:
        doubt = r["Doubt"]
        rid = r["ID"]
        if not rid:
            doubt = "; ".join(x for x in ("the printed ID could not be read: type it from the paper", doubt) if x)
        rows.append({"ID": rid, "HOURS": r["HOURS"], "OT": r["OT"] or "", "Doubt": doubt})
    notcame = [{"ID": i or f"(row {no}, no ID read)", "Name": why} for no, i, why in skipped]
    info = {"date_text": date_text, "team": team, "night": night,
            "problems": ("I turned / sharpened this photo (" + str((res.get("info") or {}).get("rescued")) + ") to read it. " if (res.get("info") or {}).get("rescued") else "") + "Read by the free reader on this laptop. Every row with a doubt is marked: check it on the paper. "
                        "Rows with no tick at all are listed under 'not counted': check they are really absent."
                        + (" ⚠️ The column headings (8H, 0.5H ...) were hard to read in this photo, so I placed the columns from the lines of the table. Look at EVERY tick with extra care."
                           if (res.get("info") or {}).get("grid_guess") else "")}
    return {"rows": rows, "not_came": notcame, "info": info, "tokens_in": 0, "tokens_out": 0}
