"""
Sync Storyblok layout (header and/or footer) from a source market to target markets.

- Header: copies the structure of the header_source market (default DK).
- Footer: copies the structure of the footer_source market (default DE).
- Every story link is re-pointed to the target market's own page (by slug).
- Labels come from config/header_labels.json and config/footer_labels.json.
  '@category_title' pulls the 'title' field from the linked product category page.
- Footer keeps the target market's own newsletter form, copyright and sign-up texts.

Modes: dry-run (no writes), draft (save, don't publish), publish.
Outputs: review .xlsx + JSON backups of every layout before it's changed.
"""
import argparse, copy, json, os, sys, time, uuid
import requests
from openpyxl import Workbook
from openpyxl.styles import Font

HERE = os.path.dirname(os.path.abspath(__file__))
cfg = lambda n: json.load(open(os.path.join(HERE, "config", n), encoding="utf-8"))
S = cfg("settings.json")
BASE = f"{S['api_base']}/spaces/{S['space_id']}"


class SB:
    def __init__(self, token):
        self.h = {"Authorization": token}

    def _req(self, method, path, **kw):
        for attempt in range(6):
            r = requests.request(method, BASE + path, headers={**self.h, "Content-Type": "application/json"}, timeout=30, **kw)
            if r.status_code == 429:
                time.sleep(1.5 * (attempt + 1))
                continue
            r.raise_for_status()
            time.sleep(0.15)
            return r.json()
        raise RuntimeError(f"Rate limited: {method} {path}")

    def get(self, path, **params):
        return self._req("GET", path, params=params)

    def put(self, path, body):
        return self._req("PUT", path, json=body)

    def all_stories(self, starts_with):
        out, page = [], 1
        while True:
            batch = self.get("/stories/", page=page, per_page=100, starts_with=starts_with)["stories"]
            if not batch:
                return out
            out += batch
            page += 1


def folder(m):
    return S["folder_pattern"].format(market=m)


class Market:
    """Slug -> page index for one market."""
    def __init__(self, sb, m):
        self.m, self.sb = m, sb
        self.idx = {s["full_slug"]: s for s in sb.all_stories(folder(m) + "/") if not s["is_folder"]}

    def find(self, path):
        for cand in (path, path.rstrip("/") + "/", path.rstrip("/")):
            if cand in self.idx:
                return self.idx[cand]
        return None

    def layout(self):
        s = self.find(f"{folder(self.m)}/{S['layout_slug']}")
        if not s:
            raise SystemExit(f"[{self.m}] layout story not found")
        return self.sb.get(f"/stories/{s['id']}")["story"]


def rule_map(rules, m):
    remap = {**rules.get("remap", {}).get("*", {}), **rules.get("remap", {}).get(m, {})}
    drop = rules.get("drop", {}).get("*", []) + rules.get("drop", {}).get(m, [])
    return remap, drop


def drop_items(items, drop, src_prefix, key="link"):
    keep = []
    for it in items:
        url = (it.get(key) or {}).get("cached_url", "")
        rel = url[len(src_prefix) + 1:] if url.startswith(src_prefix) else url
        if any(d.strip("/") in rel for d in drop):
            continue
        keep.append(it)
    return keep


def relink(block, src, tgt, remap, problems, review, part, label_hint=""):
    """Re-point every story multilink in block from src market to tgt market."""
    sp, tp = folder(src.m), folder(tgt.m)

    def walk(o, lab):
        if isinstance(o, dict):
            if "component" in o:
                o["_uid"] = str(uuid.uuid4())
            lab = o.get("title") or o.get("Title") or o.get("text") or lab
            for k, v in o.items():
                if isinstance(v, dict) and v.get("fieldtype") == "multilink" and v.get("linktype") == "story":
                    url = v.get("cached_url", "")
                    if not url.startswith(sp):
                        continue
                    new = tp + url[len(sp):]
                    for a, b in remap.items():
                        new = new.replace(a, b)
                    page = tgt.find(new)
                    if not page:
                        problems.append(f"[{tgt.m}] {part}: no page for '{lab}' -> {new}")
                        review.append([tgt.m, part, lab, k, new, "MISSING PAGE"])
                        continue
                    v["id"], v["cached_url"] = page["uuid"], new
                    review.append([tgt.m, part, lab, k, new, page["uuid"]])
                elif k != "form":
                    walk(v, lab)
        elif isinstance(o, list):
            for x in o:
                walk(x, lab)

    walk(block, label_hint)


def category_title(sb, tgt, url):
    page = tgt.find(url)
    return sb.get(f"/stories/{page['id']}")["story"]["content"].get("title", "").strip() if page else None


def build_header(sb, src, src_layout, tgt, problems, review):
    labels = cfg("header_labels.json").get(tgt.m)
    if not labels:
        raise SystemExit(f"[{tgt.m}] no header labels in config/header_labels.json")
    remap, drop = rule_map(cfg("link_rules.json")["header"], tgt.m)
    h = copy.deepcopy(src_layout["content"]["header"])
    nav = h[0]["newLinks"]
    if len(nav) != len(labels["menu"]):
        raise SystemExit(f"[{tgt.m}] header: source has {len(nav)} menu items, labels have {len(labels['menu'])}")
    for item, t in zip(nav, labels["menu"]):
        item["Title"] = t
    cols = [c for item in nav for c in item.get("columns", [])]
    if len(cols) != len(labels["columns"]):
        raise SystemExit(f"[{tgt.m}] header: source has {len(cols)} columns, labels have {len(labels['columns'])}")
    need_cat = []
    for col, lab in zip(cols, labels["columns"]):
        col["links"] = drop_items(col["links"], drop, folder(src.m))
        if len(col["links"]) != len(lab["links"]):
            raise SystemExit(f"[{tgt.m}] header column '{col['title']}': {len(col['links'])} links in source, {len(lab['links'])} labels")
        col["title"], col["viewAllLabel"] = lab["title"], lab["viewAll"]
        for ln, t in zip(col["links"], lab["links"]):
            if t == "@category_title":
                need_cat.append(ln)
            else:
                ln["title"] = t
    relink(h, src, tgt, remap, problems, review, "header")
    for ln in need_cat:  # after relink, so the URL is the target market's page
        ct = category_title(sb, tgt, ln["link"]["cached_url"])
        if ct:
            ln["title"] = ct
        else:
            problems.append(f"[{tgt.m}] header: no category title for {ln['link']['cached_url']}")
    return h


def build_footer(src, src_layout, tgt, tgt_layout, problems, review):
    labels = cfg("footer_labels.json").get(tgt.m)
    if not labels:
        raise SystemExit(f"[{tgt.m}] no footer labels in config/footer_labels.json")
    remap, drop = rule_map(cfg("link_rules.json")["footer"], tgt.m)
    f = copy.deepcopy(src_layout["content"]["footer"][0])
    cur = tgt_layout["content"]["footer"][0]
    for k in S["footer_keep_from_target"] + ["_uid"]:
        if k in cur:
            f[k] = copy.deepcopy(cur[k])
    f["contactUs"] = labels["contactUs"]
    cats = f["categories"]
    if len(cats) != len(labels["categories"]):
        raise SystemExit(f"[{tgt.m}] footer: source has {len(cats)} categories, labels have {len(labels['categories'])}")
    for c, lab in zip(cats, labels["categories"]):
        c["links"] = drop_items(c["links"], drop, folder(src.m))
        if len(c["links"]) != len(lab["links"]):
            raise SystemExit(f"[{tgt.m}] footer category '{c['title']}': {len(c['links'])} links in source, {len(lab['links'])} labels")
        c["title"] = lab["title"]
        for ln, t in zip(c["links"], lab["links"]):
            ln["title"] = t
    f["legalContent"] = drop_items(f["legalContent"], drop, folder(src.m))
    if len(f["legalContent"]) != len(labels["legal"]):
        raise SystemExit(f"[{tgt.m}] footer legal: {len(f['legalContent'])} links in source, {len(labels['legal'])} labels")
    for ln, t in zip(f["legalContent"], labels["legal"]):
        ln["title"] = t
    relink(f, src, tgt, remap, problems, review, "footer")
    return [f]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--parts", default="header,footer", help="header, footer or header,footer")
    ap.add_argument("--markets", default="dk,no,se,fi")
    ap.add_argument("--mode", choices=["dry-run", "draft", "publish"], default="dry-run")
    ap.add_argument("--out", default="output")
    a = ap.parse_args()
    token = os.environ.get("STORYBLOK_TOKEN")
    if not token:
        raise SystemExit("Set STORYBLOK_TOKEN")
    parts = [p.strip() for p in a.parts.split(",") if p.strip()]
    markets = [m.strip().lower() for m in a.markets.split(",") if m.strip()]
    os.makedirs(os.path.join(a.out, "backups"), exist_ok=True)
    sb = SB(token)
    cache = {}
    M = lambda m: cache.setdefault(m, Market(sb, m))
    review, problems, results = [], [], []

    for m in markets:
        tgt = M(m)
        layout = tgt.layout()
        json.dump(layout, open(os.path.join(a.out, "backups", f"layout_{m}_before.json"), "w"), ensure_ascii=False, indent=1)
        content = dict(layout["content"])
        changed = []
        n_problems = len(problems)
        if "header" in parts:
            if m == S["header_source"]:
                print(f"[{m}] skip header (this is the header source)")
            else:
                src = M(S["header_source"])
                content["header"] = build_header(sb, src, src.layout(), tgt, problems, review)
                changed.append("header")
        if "footer" in parts:
            if m == S["footer_source"]:
                print(f"[{m}] skip footer (this is the footer source)")
            else:
                src = M(S["footer_source"])
                content["footer"] = build_footer(src, src.layout(), tgt, layout, problems, review)
                changed.append("footer")
        if not changed:
            continue
        if len(problems) > n_problems:
            results.append([m, ",".join(changed), "NOT WRITTEN - missing pages, see problems"])
            continue
        if a.mode == "dry-run":
            results.append([m, ",".join(changed), "dry-run (not written)"])
            continue
        body = {"story": {"content": content}}
        if a.mode == "publish":
            body["publish"] = 1
        sb.put(f"/stories/{layout['id']}", body)
        after = sb.get(f"/stories/{layout['id']}")["story"]
        ok = all(after["content"].get(p) == content[p] for p in changed)
        results.append([m, ",".join(changed), f"{a.mode} {'OK' if ok else 'VERIFY FAILED'}"])

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.append(["Market", "Parts", "Result"])
    for r in results:
        ws.append(r)
    ws2 = wb.create_sheet("Links")
    ws2.append(["Market", "Part", "Label", "Field", "Target page", "Page UUID"])
    for r in review:
        ws2.append(r)
    ws3 = wb.create_sheet("Problems")
    ws3.append(["Problem"])
    for p in problems:
        ws3.append([p])
    for w in (ws, ws2, ws3):
        for c in w[1]:
            c.font = Font(bold=True)
        for col in w.columns:
            w.column_dimensions[col[0].column_letter].width = min(80, max(12, max(len(str(c.value or "")) for c in col) + 2))
    path = os.path.join(a.out, "layout_sync_review.xlsx")
    wb.save(path)

    print("\nResults:")
    for r in results:
        print("  ", " | ".join(r))
    if problems:
        print("\nProblems:")
        for p in problems:
            print("  ", p)
    print(f"\nReview file: {path}")
    with open(os.path.join(a.out, "summary.txt"), "w") as fh:
        fh.write("\n".join(" | ".join(r) for r in results) + ("\n\nProblems:\n" + "\n".join(problems) if problems else ""))
    if problems:
        sys.exit(2)


if __name__ == "__main__":
    main()
