"""
DSR Second Brain — Text-based Parser
=====================================
Parses DSR data from SharePoint MCP text extraction (not raw xlsx).
Used when binary Excel download is unavailable.

Usage:
  python process_dsr_text.py raw/<date>.txt

Appends snapshot to knowledge/dsr_history.json and re-injects dashboard.
"""

import sys, json, os, re
from datetime import datetime

BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
KNOWLEDGE    = os.path.join(BASE_DIR, "knowledge", "dsr_history.json")
TARGETS_FILE = os.path.join(BASE_DIR, "knowledge", "targets.json")
DASHBOARD    = os.path.join(BASE_DIR, "DSR_Dashboard.html")

MONTHLY_TARGET = 819086086.8313   # May 2026 grand total target from Section VII

def load_knowledge():
    if not os.path.exists(KNOWLEDGE):
        return {"_meta": {}, "snapshots": []}
    with open(KNOWLEDGE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_knowledge(kb):
    with open(KNOWLEDGE, "w", encoding="utf-8") as f:
        json.dump(kb, f, indent=2, ensure_ascii=False)

def load_targets():
    if not os.path.exists(TARGETS_FILE):
        return {}
    with open(TARGETS_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
    return {k: v for k, v in data.items() if not k.startswith("_")}

def M(s):
    """Convert '94M' or '94.2M' or '0.1M' string to float."""
    m = re.match(r'([\d.]+)M', str(s).strip())
    return float(m.group(1)) * 1_000_000 if m else 0.0

def parse_dsr_text(text, filename):
    snap = {
        "date": "", "shortDate": "",
        "processedAt": datetime.utcnow().isoformat() + "Z",
        "sourceFile": os.path.basename(filename),
        "s2Total": {}, "s1Total": {}, "s4aTotal": {}, "s7Total": {},
        "projects": [], "salesGroups": [],
        "etsTermsTotal": {}, "htsTermsTotal": {}, "pdRSTotal": {},
    }

    # ── Date ──────────────────────────────────────────────────────────────────
    dm = re.search(r'as of (\w+ \d{1,2},?\s*\d{4})', text, re.IGNORECASE)
    if dm:
        snap["date"] = dm.group(1).strip().replace(',', ',')
        m2 = re.search(r'(\w+ \d{1,2})', snap["date"])
        snap["shortDate"] = m2.group(1) if m2 else snap["date"]

    # ── Grand Total RS (Section II) ────────────────────────────────────────
    # Pattern: "Grand Total   N   XM   0   0M   N   XM   N   XM   0   0M"
    gt = re.search(
        r'Grand Total\s+(\d+)\s+([\d.]+M)\s+\d+\s+[\d.]+M\s+\d+\s+[\d.]+M\s+(\d+)\s+([\d.]+M)',
        text
    )
    rs_u, rs_ncp, ps_slots, ps_ncp = 0, 0.0, 0, 0.0
    if gt:
        rs_u     = int(gt.group(1))
        rs_ncp   = M(gt.group(2))
        ps_slots = int(gt.group(3))
        ps_ncp   = M(gt.group(4))

    # Try to find exact RS NCP from the unit-level total line at bottom of text
    # This appears as a large bare number near the end (e.g., 94,194,132)
    if rs_ncp > 0:
        for cand in re.findall(r'(?<![.\d])([\d]{2,3},[\d]{3},[\d]{3}(?:\.[\d]+)?)(?![.\d])', text):
            val = float(cand.replace(',', ''))
            if abs(val - rs_ncp) < 2_000_000 and val > 5_000_000:
                rs_ncp = val
                break

    snap["s2Total"] = {"rsU": rs_u, "rsNCP": rs_ncp, "psSlots": ps_slots, "psNCP": ps_ncp}
    snap["s1Total"] = {"rsU": rs_u, "rsNCP": rs_ncp, "totU": rs_u, "totNCP": rs_ncp,
                       "psSlots": ps_slots, "psNCP": ps_ncp}

    # ── Project groups ────────────────────────────────────────────────────────
    for proj in ["ETS", "HTS", "New Projects"]:
        pm = re.search(rf'(?<!\w){re.escape(proj)}\s+(\d+)\s+([\d.]+M)', text)
        if pm:
            snap["projects"].append({
                "name": proj, "isGroup": True,
                "rsU": int(pm.group(1)), "rsNCP": M(pm.group(2))
            })

    # ── Local vs International from unit-level sales channel column ───────────
    local_matches = re.findall(r'Local \((?:Digital|Non-Digital)\)', text)
    intl_matches  = re.findall(r'International \((?:Digital|Non-Digital)\)', text)
    loc_u = len(local_matches)
    int_u = len(intl_matches)
    total_found = loc_u + int_u
    if total_found > 0 and total_found == rs_u:
        loc_frac  = loc_u / total_found
        snap["s4aTotal"] = {
            "localU": loc_u,  "localNCP": round(rs_ncp * loc_frac),
            "intlU":  int_u,  "intlNCP":  round(rs_ncp * (1 - loc_frac)),
            "totU":   rs_u,   "totNCP":   rs_ncp,
        }
    else:
        # Fallback: use ratio from existing snapshots (~25% local for May 2026)
        loc_frac = 0.25
        snap["s4aTotal"] = {
            "localU": round(rs_u * loc_frac), "localNCP": round(rs_ncp * loc_frac),
            "intlU":  round(rs_u * (1 - loc_frac)), "intlNCP":  round(rs_ncp * (1 - loc_frac)),
            "totU":   rs_u, "totNCP": rs_ncp,
        }

    # ── Performance ───────────────────────────────────────────────────────────
    snap["s7Total"] = {
        "tgt":  MONTHLY_TARGET,
        "act":  rs_ncp,
        "perf": rs_ncp / MONTHLY_TARGET if MONTHLY_TARGET > 0 else 0,
    }

    # ── Payment terms totals (approximate from project totals) ────────────────
    ets    = next((p for p in snap["projects"] if p["name"] == "ETS"), {})
    hts    = next((p for p in snap["projects"] if p["name"] == "HTS"), {})
    np_    = next((p for p in snap["projects"] if p["name"] == "New Projects"), {})
    snap["etsTermsTotal"] = {"u": ets.get("rsU", 0),  "ncp": ets.get("rsNCP", 0)}
    snap["htsTermsTotal"] = {
        "u":   hts.get("rsU", 0) + np_.get("rsU", 0),
        "ncp": hts.get("rsNCP", 0) + np_.get("rsNCP", 0),
    }
    snap["pdRSTotal"] = {
        "totU": rs_u, "totNCP": rs_ncp, "psU": ps_slots, "psNCP": ps_ncp
    }

    return snap


def inject_block(html, start_marker, end_marker, new_content):
    pattern = rf'{re.escape(start_marker)}.*?{re.escape(end_marker)}'
    new_block = f"{start_marker}\n{new_content}\n{end_marker}"
    return re.sub(pattern, new_block, html, flags=re.DOTALL)


def update_dashboard(kb, targets):
    if not os.path.exists(DASHBOARD):
        return
    with open(DASHBOARD, "r", encoding="utf-8") as f:
        html = f.read()

    history_json = json.dumps(kb["snapshots"], ensure_ascii=False)
    targets_json = json.dumps(targets, ensure_ascii=False)

    html = inject_block(html,
        "<!--DSR_HISTORY_BLOCK_START-->", "<!--DSR_HISTORY_BLOCK_END-->",
        f"<script>window.DSR_HISTORY = {history_json};</script>"
    )
    html = inject_block(html,
        "<!--DSR_TARGETS_BLOCK_START-->", "<!--DSR_TARGETS_BLOCK_END-->",
        f"<script>window.DSR_TARGETS = {targets_json};</script>"
    )
    with open(DASHBOARD, "w", encoding="utf-8") as f:
        f.write(html)

    # Also copy as index.html for GitHub Pages
    import shutil
    try:
        shutil.copy2(DASHBOARD, os.path.join(BASE_DIR, "index.html"))
    except Exception:
        pass
    print(f"  Dashboard updated with {len(kb['snapshots'])} snapshot(s).")


def main():
    if len(sys.argv) < 2:
        print("Usage: python process_dsr_text.py <path_to_text_file>")
        sys.exit(1)

    filepath = sys.argv[1]
    if not os.path.exists(filepath):
        print(f"File not found: {filepath}")
        sys.exit(1)

    with open(filepath, "r", encoding="utf-8") as f:
        text = f.read()

    print(f"\nParsing: {os.path.basename(filepath)}")
    snap = parse_dsr_text(text, filepath)
    print(f"  Date: {snap['date']}")
    print(f"  RS: {snap['s2Total'].get('rsU', 0)} units / ~{round(snap['s2Total'].get('rsNCP', 0)/1e6)}M")

    kb = load_knowledge()
    existing = [s["shortDate"] for s in kb["snapshots"]]
    if snap["shortDate"] in existing:
        print(f"  Updating existing snapshot for {snap['shortDate']}")
        kb["snapshots"] = [s for s in kb["snapshots"] if s["shortDate"] != snap["shortDate"]]

    kb["snapshots"].append(snap)
    kb["snapshots"].sort(key=lambda s: s.get("date", ""))
    save_knowledge(kb)
    print(f"  Knowledge base: {len(kb['snapshots'])} total snapshot(s).")

    targets = load_targets()
    update_dashboard(kb, targets)
    print("Done.\n")


if __name__ == "__main__":
    main()
