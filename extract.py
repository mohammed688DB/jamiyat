import openpyxl, sys, json, re, datetime, warnings
warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding='utf-8')

AZA = r"C:\Users\Dell\Desktop\جمعية العزاء.xlsx"
MAJ = r"C:\Users\Dell\Desktop\جمعية المجلس.xlsx"
HERE = __import__("os").path.dirname(__import__("os").path.abspath(__file__))
OUT = __import__("os").path.join(HERE, "data.json")


def clean(s):
    return re.sub(r"\s+", " ", str(s).replace("ـ", "")).strip() if s is not None else ""


def key(name):
    s = clean(name)
    s = re.sub("[أإآ]", "ا", s).replace("ى", "ي").replace("ة", "ه")
    s = re.sub(r"\bين\b", "بن", s)
    return " ".join(t for t in s.split() if t != "بن")


def num(v):
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip())
    except Exception:
        return 0.0


def fmt_date(v):
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d")
    if v is None:
        return ""
    s = clean(v)
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", s)
    if m:
        return f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
    return s  # e.g. "2016" or "عام 2014"


class Registry:
    def __init__(self, prefix):
        self.prefix, self.groups, self.members, self.by_key, self.gid_by_name = prefix, [], [], {}, {}

    def group(self, name):
        name = clean(name)
        k = key(name.replace("مجموعة", ""))
        if k not in self.gid_by_name:
            gid = f"{self.prefix}g{len(self.groups) + 1}"
            self.groups.append({"id": gid, "name": name})
            self.gid_by_name[k] = gid
        return self.gid_by_name[k]

    def member(self, name, gid, **extra):
        k = key(name)
        if k in self.by_key:
            m = self.by_key[k]
            if gid and m["groupId"] is None:
                m["groupId"] = gid
            for f, v in extra.items():
                if v and not m.get(f):
                    m[f] = v
            return m["id"]
        mid = f"{self.prefix}m{len(self.members) + 1}"
        m = {"id": mid, "name": clean(name), "groupId": gid, "phone": "", "note": "", "active": True}
        m.update({f: v for f, v in extra.items() if v})
        self.members.append(m)
        self.by_key[k] = m
        return mid


def is_member_row(a, b):
    b = clean(b)
    return b and b.startswith(("مجموعة", "الأس", "المجم", "ت")) is False and not (isinstance(a, str) and a.strip() in ("ت", "المجموع", "المجموع الكلي")) and len(b) > 5


# ---------------------------------------------------------------- جمعية العزاء
wb = openpyxl.load_workbook(AZA, data_only=True)
reg = Registry("a")
years, checks = {}, {}
sheets = [(int(re.search(r"\d{4}", ws.title).group()), ws) for ws in wb.worksheets if re.search(r"عام\s*\d{4}", ws.title)]
# process newest first so group assignments come from the grouped (recent) sheets
for year, ws in sorted(sheets, key=lambda t: -t[0]):
    hdr = [clean(ws.cell(r, 3).value) for r in range(1, 5)]
    cur_gid, rows, sheet_total = None, {}, 0.0
    phone_col = None
    for r in range(1, 5):
        for c in range(15, 18):
            if "الهات" in clean(ws.cell(r, c).value):
                phone_col = c
    for r in range(1, ws.max_row + 1):
        a, b = ws.cell(r, 1).value, ws.cell(r, 2).value
        if isinstance(a, str) and clean(a).startswith("مجموعة"):
            cur_gid = reg.group(a)
            continue
        if isinstance(a, str) and clean(a).startswith("المجموع"):
            if clean(a) == "المجموع الكلي" or year < 2020:
                sheet_total = sum(num(ws.cell(r, c).value) for c in range(3, 15))
            continue
        if isinstance(b, str) and clean(b).startswith("المجموع"):
            sheet_total = sum(num(ws.cell(r, c).value) for c in range(3, 15))
            continue
        if not isinstance(b, str) or not is_member_row(a, b):
            continue
        name = clean(b)
        donor = "فاعل خير" in name
        phone = clean(ws.cell(r, phone_col).value) if phone_col else ""
        gid = cur_gid
        mid = reg.member(name, gid, phone=phone, exempt=donor)
        months = [num(ws.cell(r, c).value) for c in range(3, 15)]
        if mid in rows:  # duplicate row in same sheet -> merge
            rows[mid] = [x + y for x, y in zip(rows[mid], months)]
        else:
            rows[mid] = months
    years[str(year)] = {"pay": rows}
    got = sum(sum(v) for v in rows.values())
    checks[year] = (got, sheet_total)

# members that only appear in the ungrouped years
old_gid = None
for m in reg.members:
    if m["groupId"] is None:
        if old_gid is None:
            old_gid = reg.group("أعضاء سابقون / غير مصنفين")
        m["groupId"] = old_gid
    if m.get("exempt"):
        m["groupId"] = reg.group("فاعل خير / تبرعات")

# members absent from the newest year are marked inactive
latest = max(int(y) for y in years)
for m in reg.members:
    m["active"] = m["id"] in years[str(latest)]["pay"]

print("=== العزاء: year totals (extracted vs sheet)")
for y in sorted(checks):
    g, t = checks[y]
    print(y, g, t, "OK" if abs(g - t) < 0.01 else "DIFF")

ws = wb["مصروفات الجمعية"]
aza_exp = []
for r in range(3, ws.max_row + 1):
    title, amt = clean(ws.cell(r, 2).value), ws.cell(r, 4).value
    if title and title != "المجموع" and amt is not None:
        aza_exp.append({"id": f"ae{len(aza_exp) + 1}", "date": fmt_date(ws.cell(r, 3).value), "title": title,
                        "amount": num(amt), "type": clean(ws.cell(r, 5).value), "reason": clean(ws.cell(r, 6).value)})
print("expenses", len(aza_exp), sum(e["amount"] for e in aza_exp))

aza = {
    "settings": {"title": "جمعية العزاء", "monthlyFee": 2, "bankBalance": 2279, "currency": "ر.ع"},
    "groups": [g for g in reg.groups if any(m["groupId"] == g["id"] for m in reg.members)],
    "members": reg.members, "years": years,
    "expenses": aza_exp, "income": [],
}
print("groups", len(reg.groups), "members", len(reg.members), "active", sum(m["active"] for m in reg.members))

# ---------------------------------------------------------------- جمعية المجلس
wb = openpyxl.load_workbook(MAJ, data_only=True)
reg = Registry("m")
MONTHS = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"]
plans = [
    {"id": "p1", "name": "الجمعية الأولى", "amount": 240, "note": "20 ريال شهرياً لمدة 12 شهر"},
    {"id": "p2", "name": "تعزيز الجمعية الأولى", "amount": 36, "note": ""},
    {"id": "p3", "name": "الجمعية الثانية", "amount": 200, "note": ""},
    {"id": "p4", "name": "الجمعية الثالثة", "amount": 100, "installment": 20, "note": "100 ريال على 5 دفعات، كل دفعة 20 ريال"},
]
payments, overrides = [], {}


def pay(mid, pid, amount, note):
    if amount:
        payments.append({"id": f"mp{len(payments) + 1}", "memberId": mid, "planId": pid, "amount": amount, "date": "", "note": note})


def walk(ws, fn, name_col=2):
    gid = None
    for r in range(1, ws.max_row + 1):
        a, b = ws.cell(r, 1).value, ws.cell(r, 2).value
        for v in (a, b):
            if isinstance(v, str) and clean(v).startswith("مجموعة"):
                gid = reg.group(v)
        if isinstance(b, str) and is_member_row(a, b) and not clean(b).startswith("مجموعة"):
            fn(r, gid, b)


# الشاملة first: defines groups/order
ws = wb["جمعية المجلس الشاملة"]
walk(ws, lambda r, gid, b: reg.member(b, gid))
# الأولى: monthly payments + تعزيز
ws = wb["جمعية المجلس الأولى"]
first_total = [0.0]


def first(r, gid, b):
    mid = reg.member(b, gid)
    for i in range(12):
        v = num(ws.cell(r, 3 + i).value)
        pay(mid, "p1", v, MONTHS[i])
        first_total[0] += v
    v = num(ws.cell(r, 15).value)
    pay(mid, "p2", v, "")
    first_total[0] += v


walk(ws, first)
ws2 = wb["جمعية المجلس الثانية"]


def second(r, gid, b):
    mid = reg.member(b, gid)
    if ws2.cell(r, 3).value in (0, None):
        overrides.setdefault(mid, {})["p3"] = 0
    pay(mid, "p3", num(ws2.cell(r, 4).value), "")


walk(ws2, second)
# الشاملة: per-member required overrides (الأولى = 0 for new group)
ws = wb["جمعية المجلس الشاملة"]


def shamila(r, gid, b):
    mid = reg.member(b, gid)
    if ws.cell(r, 3).value == 0:
        overrides.setdefault(mid, {})["p1"] = 0


walk(ws, shamila)
# قائمة التعزيز: 5 installments
ws = wb["قايمة التعزيز"]
for r in range(3, ws.max_row + 1):
    b = ws.cell(r, 2).value
    if isinstance(b, str) and is_member_row(ws.cell(r, 1).value, b):
        mid = reg.member(b, None)
        for i, lbl in enumerate(["الدفعة الأولى", "الدفعة الثانية", "الدفعة الثالثة", "الدفعة الرابعة", "الدفعة الخامسة"]):
            pay(mid, "p4", num(ws.cell(r, 3 + i).value), lbl)
# notes from defaulters list
ws = wb["قائمة المتخلفين عن الدفع"]
for r in range(4, ws.max_row + 1):
    b, note = ws.cell(r, 2).value, clean(ws.cell(r, 5).value)
    if isinstance(b, str) and note and key(b) in reg.by_key:
        reg.by_key[key(b)]["note"] = note

# building expenses
ws = wb["مصروفات بناء المجلس"]
maj_exp = []
for r in range(3, 31):
    title, amt = clean(ws.cell(r, 2).value), ws.cell(r, 3).value
    if title and amt is not None:
        maj_exp.append({"id": f"me{len(maj_exp) + 1}", "date": "", "title": title, "amount": num(amt), "type": "بناء المجلس", "reason": ""})
stages = []
for r in range(3, 13):
    t, a = clean(ws.cell(r, 9).value), ws.cell(r, 10).value
    if t:
        stages.append({"id": f"st{len(stages) + 1}", "title": t, "amount": num(a), "done": False})
cpay = []
for r in range(3, 17):
    a = ws.cell(r, 6).value
    if isinstance(a, (int, float)) and a:
        cpay.append({"id": f"cp{len(cpay) + 1}", "date": "", "title": f"دفعة رقم {len(cpay) + 1}", "amount": num(a)})
print("contractor paid", sum(c["amount"] for c in cpay))
maj_income = [
    {"id": "mi1", "date": "", "title": "بيع سابة الجل", "amount": 1450},
    {"id": "mi2", "date": "", "title": "شركة الألمنيوم", "amount": 2000},
]
used = {m["groupId"] for m in reg.members}
maj = {
    "settings": {"title": "جمعية المجلس", "currency": "ر.ع", "contractTotal": 22500, "contractor": "المقاول"},
    "groups": [g for g in reg.groups if g["id"] in used], "members": reg.members, "plans": plans, "overrides": overrides,
    "payments": payments, "expenses": maj_exp, "income": maj_income, "stages": stages, "contractorPayments": cpay,
}
print("=== المجلس groups", len(reg.groups), "members", len(reg.members), "payments", len(payments))
for p in plans:
    print(p["name"], sum(x["amount"] for x in payments if x["planId"] == p["id"]))
print("first sheet total", first_total[0])
print("expenses", len(maj_exp), sum(e["amount"] for e in maj_exp))
print("overrides", {reg.members[int(k[2:]) - 1]["name"]: v for k, v in overrides.items()})
for g in reg.groups:
    print(g["id"], g["name"], [m["name"] for m in reg.members if m["groupId"] == g["id"]].__len__())
print("no group:", [m["name"] for m in reg.members if not m["groupId"]])

# ---------------------------------------------------------------- جمعية مغلقة (closed rotating association)
ROSCA = r"C:\Users\Dell\Desktop\جمعية مغلقةمعدلة.xlsx"
wb = openpyxl.load_workbook(ROSCA, data_only=True)
ws = next(w for w in wb.worksheets if "حسب عدد الإشتراك" in w.title)
subs = []
for r in range(5, ws.max_row + 1):
    a, b, c = ws.cell(r, 1).value, ws.cell(r, 2).value, ws.cell(r, 3).value
    if isinstance(a, (int, float)) and isinstance(b, str) and b.strip() and c:
        subs.append({"id": f"r_m1_s{len(subs) + 1}", "order": len(subs) + 1, "name": clean(b), "inst": num(c),
                     "phone": "", "note": "", "payout": 0, "active": True})
rosca = {
    "settings": {"title": "الجمعية المغلقة", "start": "2026-06", "months": 24, "currency": "ر.ع", "treasurer": ""},
    "settingsTs": 1, "members": subs, "pay": {s["id"]: [0] * 24 for s in subs}, "payTs": {}, "receipts": [], "groups": [],
}
print("=== جمعية مغلقة:", len(subs), "subscriptions, monthly pot", sum(s["inst"] for s in subs))

data = {"version": 1, "aza": aza, "majlis": maj, "roscaIds": ["r_m1"], "r_m1": rosca}
json.dump(data, open(OUT, "w", encoding="utf-8"), ensure_ascii=False)
print("written", OUT)

import os
TPL = os.path.join(HERE, "app-template.html")
DEST_DIR = r"C:\Users\Dell\Desktop\تطبيق الجمعيات الأهلية"
if os.path.exists(TPL):
    os.makedirs(DEST_DIR, exist_ok=True)
    html = open(TPL, encoding="utf-8").read()
    payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    html = html.replace("/*__SEED__*/null", payload)
    dest = os.path.join(DEST_DIR, "الجمعيات الأهلية.html")
    open(dest, "w", encoding="utf-8").write(html)
    print("built", dest)


def build_pwa(tpl, out):
    import hashlib
    os.makedirs(out, exist_ok=True)
    open(os.path.join(out, "index.html"), "w", encoding="utf-8").write(tpl)
    ver = hashlib.sha1(tpl.encode("utf-8")).hexdigest()[:10]
    manifest = {
        "name": "الجمعيات الأهلية", "short_name": "الجمعيات", "lang": "ar", "dir": "rtl",
        "start_url": "./", "scope": "./", "display": "standalone",
        "background_color": "#f5f4f0", "theme_color": "#1f6f6b",
        "icons": [{"src": "icon-192.png", "sizes": "192x192", "type": "image/png"},
                  {"src": "icon-512.png", "sizes": "512x512", "type": "image/png"},
                  {"src": "icon-512.png", "sizes": "512x512", "type": "image/png", "purpose": "maskable"}],
        # lets WhatsApp's "share" send a sync file straight into the installed app (Android)
        "share_target": {
            "action": "./share-target", "method": "POST", "enctype": "multipart/form-data",
            "params": {"files": [{"name": "file", "accept": ["application/json", "text/plain", "application/octet-stream", ".json", "*/*"]}]},
        },
    }
    json.dump(manifest, open(os.path.join(out, "manifest.webmanifest"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    sw = """const CACHE = "ahliya-%s";
const CORE = ["./", "./index.html", "./manifest.webmanifest", "./icon-192.png", "./icon-512.png"];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(CORE))); self.skipWaiting(); });
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE && k !== "ahliya-share").map(k => caches.delete(k)))));
  self.clients.claim();
});
// serve from cache instantly (works offline), refresh the cache in the background
self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  // a file shared to the app (WhatsApp → مشاركة → الجمعيات): park it, then open the app which merges it
  if (e.request.method === "POST" && url.pathname.endsWith("/share-target")) {
    e.respondWith((async () => {
      try {
        const form = await e.request.formData(), file = form.get("file");
        if (file) {
          const c = await caches.open("ahliya-share");
          await c.put(new URL("shared-file", self.registration.scope).href,
            new Response(file, { headers: { "x-name": encodeURIComponent(file.name || "") } }));
        }
      } catch (err) {}
      return Response.redirect(new URL("./?shared=1#aza/sync", self.registration.scope).href, 303);
    })());
    return;
  }
  if (e.request.method !== "GET") return;
  e.respondWith(caches.match(e.request, { ignoreSearch: true }).then(hit => {
    const net = fetch(e.request).then(res => {
      if (res.ok || res.type === "opaque") { const copy = res.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); }
      return res;
    }).catch(() => hit);
    return hit || net;
  }));
});
""" % ver
    open(os.path.join(out, "sw.js"), "w", encoding="utf-8").write(sw)
    make_icons(out)
    print("built PWA", out, ver)


def make_icons(out):
    from PIL import Image, ImageDraw
    for size in (192, 512):
        s = 4  # supersample
        S = size * s
        im = Image.new("RGB", (S, S), (31, 111, 107))
        d = ImageDraw.Draw(im)
        # house (majlis)
        cx, base = S * 0.5, S * 0.74
        w, h = S * 0.46, S * 0.24
        d.rectangle([cx - w / 2, base - h, cx + w / 2, base], fill=(255, 255, 255))
        d.polygon([(cx - w / 2 - S * 0.05, base - h), (cx, base - h - S * 0.2), (cx + w / 2 + S * 0.05, base - h)], fill=(255, 255, 255))
        d.rectangle([cx - S * 0.055, base - S * 0.13, cx + S * 0.055, base], fill=(31, 111, 107))
        # crescent
        r = S * 0.09; mx, my = cx + S * 0.2, S * 0.22
        d.ellipse([mx - r, my - r, mx + r, my + r], fill=(242, 201, 120))
        d.ellipse([mx - r + S * 0.04, my - r - S * 0.02, mx + r + S * 0.04, my + r - S * 0.02], fill=(31, 111, 107))
        im.resize((size, size), Image.LANCZOS).save(os.path.join(out, f"icon-{size}.png"))


if os.path.exists(TPL):
    # phone app (PWA): same app WITHOUT embedded member data — data arrives via a sync file
    build_pwa(open(TPL, encoding="utf-8").read(), os.path.join(DEST_DIR, "تطبيق الجوال"))
