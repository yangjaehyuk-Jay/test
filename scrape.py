# -*- coding: utf-8 -*-
"""
2단계: 파생상품 정보 일괄 조회

브라우저를 띄우면 사람이 로그인한다. 로그인이 끝나면 화면 우측 하단에
'자동 조회 시작' 버튼이 활성화된다. 누르면 companies.csv 의 사업자등록번호를
차례로 조회해서 엑셀로 저장한다.

조회는 화면을 클릭하지 않고, 화면이 쓰는 것과 같은 요청을 그대로 보낸다.
    POST /bizkiR0101SelectList.do   (파생상품 조회)
    POST /bizcmEventLoggerInsert.do (사이트 감사로그. 수동 조회와 동일하게 남긴다)

실행:  python scrape.py
"""
import sys, re, csv, json, time, datetime, pathlib, argparse
import unicodedata

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from playwright.sync_api import sync_playwright
import nexacro as nx

BASE = pathlib.Path(__file__).parent
PROFILE_DIR = BASE / "browser_profile"
OUT_DIR = BASE / "output"
LOG_DIR = BASE / "logs"
COMPANIES = BASE / "companies.csv"

HOST = "https://bizinfo.kfb.or.kr"
START_URL = HOST + "/biz/index.html"
POPUP_URL = (HOST + "/biz/popup.html"
             "?formname=BIZKI%3A%3ABIZKI_T0101.xfdl"
             "&framename=KI01_BIZKI_T0101.xfdl")
SVC_SELECT = HOST + "/bizkiR0101SelectList.do"
SVC_LOG = HOST + "/bizcmEventLoggerInsert.do"
SVC_LOGIN2 = "bizliS0102Login2.do"

DELAY_SEC = 2.5        # 요청 간 간격. 서버가 연속 호출을 차단하므로 필수
RETRY = 3
RETRY_WAIT = 8

XHR_HEADERS = {
    "content-type": "text/xml",
    "accept": "application/xml, text/xml, */*",
    "x-requested-with": "XMLHttpRequest",
    "cache-control": "no-cache, no-store",
    "pragma": "no-cache",
    "referer": POPUP_URL,
}

# 응답 컬럼 -> 엑셀 열 이름.
# FX=선물환, OP=옵션, CRS=통화스왑. BAL=잔액, VLD_CNT=유효건수.
COLMAP = [
    ("사업자등록번호", "_bizno"),
    ("기업명", "KIKO_CUST_NM"),
    ("결과코드", "RSP_CD"),
    ("결과내용", "RSP_NM"),
    ("선물환_건수", "FX_VLD_CNT"),
    ("선물환_매수잔액", "FX_BUY_BAL"),
    ("선물환_매도잔액", "FX_SELL_BAL"),
    ("선물환_최종갱신일", "FX_TRN_LST_UPD_DT"),
    ("옵션_건수", "OP_VLD_CNT"),
    ("옵션_매수잔액", "OP_BUY_BAL"),
    ("옵션_매도잔액", "OP_SELL_BAL"),
    ("옵션_매수콜", "OP_BUY_CALL_BAL"),
    ("옵션_매수풋", "OP_BUY_PUT_BAL"),
    ("옵션_매도콜", "OP_SELL_CALL_BAL"),
    ("옵션_매도풋", "OP_SELL_PUT_BAL"),
    ("옵션_최종갱신일", "OP_TRN_LST_UPD_DT"),
    ("통화스왑_건수", "CRS_VLD_CNT"),
    ("통화스왑_매수잔액", "CRS_BUY_BAL"),
    ("통화스왑_매도잔액", "CRS_SELL_BAL"),
    ("통화스왑_최종갱신일", "CRS_TRN_LST_UPD_DT"),
    ("USD환산금액", "BNS_AMT_USD_CVT_AMT"),
    ("조회일시", "RETRIEVE_DTTM"),
]
NUMERIC = {c for c, _ in COLMAP if "잔액" in c or "건수" in c or "환산" in c}

INJECT_JS = """
(() => {
  if (window.top !== window || window._popup) return;
  var add = function () {
    if (!document.body || document.getElementById('__kfb_run')) return;
    var box = document.createElement('div');
    box.id = '__kfb_run';
    box.style.cssText = 'position:fixed;right:14px;bottom:14px;z-index:2147483647;'
      + 'font:13px/1.5 sans-serif;text-align:right;';
    var msg = document.createElement('div');
    msg.id = '__kfb_msg';
    msg.style.cssText = 'margin-bottom:6px;padding:6px 10px;background:#222;'
      + 'color:#fff;border-radius:4px;opacity:.9;';
    msg.textContent = '\\ub85c\\uadf8\\uc778 \\ub300\\uae30 \\uc911...';
    var btn = document.createElement('button');
    btn.id = '__kfb_btn';
    btn.disabled = true;
    btn.textContent = '\\u25b6 \\uc790\\ub3d9 \\uc870\\ud68c \\uc2dc\\uc791';
    btn.style.cssText = 'padding:11px 16px;background:#999;color:#fff;border:0;'
      + 'border-radius:5px;cursor:not-allowed;font-size:14px;'
      + 'box-shadow:0 2px 8px rgba(0,0,0,.35);';
    btn.onclick = function () {
      btn.disabled = true; btn.style.background = '#999';
      btn.style.cursor = 'not-allowed';
      window.kfbStart();
    };
    box.appendChild(msg); box.appendChild(btn);
    document.body.appendChild(box);
  };
  if (document.readyState === 'loading')
    document.addEventListener('DOMContentLoaded', add);
  else add();
  setTimeout(add, 3000); setTimeout(add, 8000);

  window.__kfbReady = function (text) {
    var b = document.getElementById('__kfb_btn');
    var m = document.getElementById('__kfb_msg');
    if (m) m.textContent = text;
    if (b) { b.disabled = false; b.style.background = '#0796ec';
             b.style.cursor = 'pointer'; }
  };
  window.__kfbMsg = function (text) {
    var m = document.getElementById('__kfb_msg');
    if (m) m.textContent = text;
  };
})();
"""


class Log:
    def __init__(self, path):
        self.fh = open(path, "w", encoding="utf-8")
        self.path = path

    def __call__(self, msg=""):
        s = str(msg)
        self.fh.write(s + "\n")
        self.fh.flush()
        try:
            print(s)
        except Exception:
            print(s.encode("ascii", "replace").decode("ascii"))

    def close(self):
        self.fh.close()


# --------------------------------------------------------------------------
def read_text_any(path):
    """엑셀/메모장이 저장한 CP949 도, UTF-8 도 읽는다."""
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "cp949", "euc-kr"):
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace"), "utf-8(손상)"


def load_companies(path):
    """companies.csv 를 읽는다. 첫 열이 사업자등록번호, 둘째 열(선택)이 메모."""
    if not path.exists():
        return []
    out = []
    text, enc = read_text_any(path)
    if enc != "utf-8-sig":
        print("  companies.csv 인코딩: %s" % enc)
    import io as _io
    if True:
        f = _io.StringIO(text)
        for row in csv.reader(f):
            if not row:
                continue
            raw = row[0].strip()
            if not raw or raw.startswith("#"):
                continue
            digits = re.sub(r"\D", "", raw)
            if len(digits) != 10:
                if digits:
                    print("  건너뜀 (10자리 아님): %s" % raw)
                continue
            memo = row[1].strip() if len(row) > 1 else ""
            out.append((digits, memo))
    # 중복 제거 (순서 유지)
    seen, uniq = set(), []
    for d, m in out:
        if d not in seen:
            seen.add(d)
            uniq.append((d, m))
    return uniq


def fmt_bizno(d):
    return "%s-%s-%s" % (d[:3], d[3:5], d[5:])


def to_num(v):
    if v is None:
        return None
    s = str(v).replace(",", "").strip()
    if s in ("", "-"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def post_xml(req_ctx, url, body, log):
    """XML POST + 재시도. 서버가 연속 호출을 차단하므로 물러섰다 다시 친다."""
    last = None
    for attempt in range(1, RETRY + 1):
        try:
            resp = req_ctx.post(url, data=body, headers=XHR_HEADERS, timeout=60000)
            text = resp.text()
            if resp.status != 200:
                last = "HTTP %s" % resp.status
            elif "<Root" not in text:
                # 차단되면 '일시적인 오류' 안내 HTML 이 온다
                last = "XML 아님 (차단 추정, %d바이트)" % len(text)
            else:
                return text
        except Exception as e:
            last = str(e)[:120]
        if attempt < RETRY:
            log("      재시도 %d/%d (%s) - %d초 대기"
                % (attempt, RETRY - 1, last, RETRY_WAIT * attempt))
            time.sleep(RETRY_WAIT * attempt)
    raise RuntimeError(last or "알 수 없는 실패")


def audit_log(req_ctx, user, act, log):
    """사이트 감사로그를 수동 조회와 동일하게 남긴다."""
    body = nx.build_request(
        {"ds_log": [{
            "USER_ID": user["USER_ID"], "USER_NM": user["USER_NM"],
            "SCREEN_ID": "/BIZKI/BIZKI_R0101",
            "SCREEN_NM": "파생상품(TAB_01)", "ACT_NM": act,
        }]},
        params={"npPfsHost": "127.0.0.1", "npPfsPort": "14440"})
    try:
        post_xml(req_ctx, SVC_LOG, body, log)
    except Exception as e:
        log("      (감사로그 실패: %s)" % e)


def query_one(req_ctx, user, bizno, st, ed, log):
    body = nx.build_request(
        {"ds_search": [{
            "USER_ID": user["USER_ID"],
            "USER_NM": user["USER_NM"],
            "KIKO_ID_NO_TP_CD": "2",              # 2 = 사업자등록번호
            "KIKO_ID_NO": "999" + bizno,          # 화면이 만드는 조회요청 식별번호
            "INQ_DV_CD": "1",
            "INQ_ST_DT": st,
            "INQ_ED_DT": ed,
            "IS_IN_USER": "N",
            "USER_ORG_CD": user["USER_ORG_CD"],
            "USER_ORG_NM": user["USER_ORG_NM"],
        }]},
        params={"npPfsHost": "127.0.0.1", "npPfsPort": "14440"})
    text = post_xml(req_ctx, SVC_SELECT, body, log)
    parsed = nx.parse_response(text)
    nx.check(parsed, text)
    etc = (parsed["datasets"].get("ds_etc") or [{}])[0]
    lst = (parsed["datasets"].get("ds_list") or [{}])[0]
    merged = dict(lst)
    merged.update({k: v for k, v in etc.items() if v not in (None, "")})
    # 기업명이 전각문자로 온다 (ＳＡＭＹＡＮＧ　ＣＯＲＰ．) -> 반각으로 정리
    nm = merged.get("KIKO_CUST_NM")
    if nm:
        merged["KIKO_CUST_NM"] = unicodedata.normalize("NFKC", nm).strip()
    return merged


PARTIAL = {"rows": [], "failed": []}


def run_batch(req_ctx, user, companies, st, ed, log, notify):
    """한 건씩 조회한다. 중간에 끊겨도 결과를 잃지 않도록 즉시 CSV 에 적는다."""
    rows, failed = PARTIAL["rows"], PARTIAL["failed"]
    total = len(companies)
    OUT_DIR.mkdir(exist_ok=True)
    part_path = OUT_DIR / ("_진행중_%s.csv"
                           % datetime.date.today().strftime("%Y%m%d"))
    cols = ["사업자등록번호", "메모"] + [c for c, _ in COLMAP if c != "_bizno"
                                        and c != "사업자등록번호"]
    part = open(part_path, "w", encoding="utf-8-sig", newline="")
    pw = csv.DictWriter(part, fieldnames=cols, extrasaction="ignore")
    pw.writeheader()

    audit_log(req_ctx, user, "화면오픈", log)
    for i, (bizno, memo) in enumerate(companies, 1):
        label = fmt_bizno(bizno) + (" (%s)" % memo if memo else "")
        log("  [%2d/%d] %s" % (i, total, label))
        notify("조회 중 %d/%d  %s" % (i, total, fmt_bizno(bizno)))
        try:
            data = query_one(req_ctx, user, bizno, st, ed, log)
            audit_log(req_ctx, user, "내역조회", log)
            rec = {"_bizno": fmt_bizno(bizno)}
            for col, key in COLMAP:
                if key == "_bizno":
                    continue
                rec[col] = data.get(key)
            rec["사업자등록번호"] = fmt_bizno(bizno)
            if memo:
                rec["메모"] = memo
            rows.append(rec)
            pw.writerow(rec); part.flush()
            log("       -> %s / %s | 선물환 %s건, 옵션 %s건, 스왑 %s건"
                % (data.get("RSP_CD"), (data.get("KIKO_CUST_NM") or "").strip(),
                   data.get("FX_VLD_CNT"), data.get("OP_VLD_CNT"),
                   data.get("CRS_VLD_CNT")))
        except Exception as e:
            log("       -> 실패: %s" % e)
            failed.append((fmt_bizno(bizno), memo, str(e)[:160]))
        if i < total:
            time.sleep(DELAY_SEC)
    part.close()
    log("  (중간저장: %s)" % part_path.name)
    return rows, failed


def write_excel(rows, failed, st, ed, log):
    import pandas as pd
    OUT_DIR.mkdir(exist_ok=True)
    today = datetime.date.today().strftime("%Y%m%d")
    path = OUT_DIR / ("파생거래_%s.xlsx" % today)

    cols = [c for c, _ in COLMAP if c != "_bizno"]
    cols = ["사업자등록번호"] + [c for c in cols if c != "사업자등록번호"]
    if any("메모" in r for r in rows):
        cols.insert(1, "메모")
    df = pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)
    for c in df.columns:
        if c in NUMERIC:
            df[c] = df[c].map(to_num)

    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="파생거래", index=False)
        meta = pd.DataFrame([
            ("조회일시", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            ("조회기간", "%s ~ %s" % (st, ed)),
            ("요청 기업수", len(rows) + len(failed)),
            ("성공", len(rows)),
            ("실패", len(failed)),
        ], columns=["항목", "값"])
        meta.to_excel(xw, sheet_name="조회정보", index=False)
        if failed:
            pd.DataFrame(failed, columns=["사업자등록번호", "메모", "사유"]) \
              .to_excel(xw, sheet_name="실패", index=False)

        ws = xw.sheets["파생거래"]
        for idx, c in enumerate(df.columns, 1):
            width = max(len(str(c)) + 2,
                        *(len(str(v)) + 2 for v in df[c].head(60).astype(str))) \
                    if len(df) else len(str(c)) + 2
            ws.column_dimensions[ws.cell(1, idx).column_letter].width = min(width, 22)
            if c in NUMERIC:
                for r in range(2, len(df) + 2):
                    ws.cell(r, idx).number_format = "#,##0.00"
        ws.freeze_panes = "B2"
    log("\n엑셀 저장 -> %s" % path)
    return path


# --------------------------------------------------------------------------
def main():
    global DELAY_SEC
    ap = argparse.ArgumentParser()
    ap.add_argument("--from", dest="st", default=None,
                    help="조회 시작일 YYYYMMDD (기본: 올해 1월 1일)")
    ap.add_argument("--to", dest="ed", default=None,
                    help="조회 종료일 YYYYMMDD (기본: 오늘)")
    ap.add_argument("--delay", type=float, default=DELAY_SEC)
    a = ap.parse_args()

    DELAY_SEC = a.delay
    today = datetime.date.today()
    st = a.st or today.replace(month=1, day=1).strftime("%Y%m%d")
    ed = a.ed or today.strftime("%Y%m%d")

    LOG_DIR.mkdir(exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    log = Log(LOG_DIR / ("scrape_" + stamp + ".log"))

    companies = load_companies(COMPANIES)
    if not companies:
        log("companies.csv 에 사업자등록번호가 없습니다: %s" % COMPANIES)
        log("한 줄에 하나씩, 10자리 숫자로 넣어주세요. (하이픈은 있어도 됩니다)")
        log.close()
        return 1
    log("조회 대상 %d개 기업 / 조회기간 %s ~ %s / 간격 %.1f초"
        % (len(companies), st, ed, DELAY_SEC))

    state = {"user": None, "start": False}

    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR), headless=False,
            locale="ko-KR", timezone_id="Asia/Seoul", viewport=None,
            args=["--disable-blink-features=AutomationControlled",
                  "--disable-popup-blocking", "--start-maximized"])
        ctx.add_init_script(INJECT_JS)
        ctx.expose_binding("kfbStart",
                           lambda source: state.__setitem__("start", True))

        def on_response(resp):
            # 로그인 응답에서 사용자/기관 정보를 주워둔다. 조회 요청에 필요하다.
            if state["user"] or SVC_LOGIN2 not in resp.url:
                return
            try:
                parsed = nx.parse_response(resp.body())
            except Exception:
                return
            for rows in parsed["datasets"].values():
                for r in rows:
                    if r.get("USER_ID") and r.get("USER_ORG_CD"):
                        state["user"] = {
                            "USER_ID": r["USER_ID"], "USER_NM": r.get("USER_NM", ""),
                            "USER_ORG_CD": r["USER_ORG_CD"],
                            "USER_ORG_NM": r.get("USER_ORG_NM", ""),
                        }
                        log("로그인 확인: %s / %s (%s)"
                            % (r["USER_ID"], r.get("USER_NM"), r.get("USER_ORG_NM")))
                        return

        ctx.on("response", on_response)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto(START_URL, wait_until="domcontentloaded")

        bar = "=" * 74
        log(bar)
        log(" 로그인하세요 (ID/PW + 추가 인증).")
        log(" 로그인이 끝나면 우측 하단 '자동 조회 시작' 버튼이 파랗게 켜집니다.")
        log(" 버튼을 누르면 %d개 기업을 차례로 조회합니다." % len(companies))
        log(bar)

        def notify(text):
            for pg in ctx.pages:
                try:
                    pg.evaluate("t => window.__kfbMsg && window.__kfbMsg(t)", text)
                except Exception:
                    pass

        armed = False
        closed = {"v": False}
        ctx.on("close", lambda _: closed.__setitem__("v", True))
        rows = failed = None
        started = time.time()
        while not closed["v"]:
            pages = [x for x in ctx.pages if not x.is_closed()]
            if not pages:
                break
            if state["user"] and not armed:
                armed = True
                txt = "준비 완료 - %d개 기업" % len(companies)
                for pg in pages:
                    try:
                        pg.evaluate("t => window.__kfbReady && window.__kfbReady(t)", txt)
                    except Exception:
                        pass
                log("버튼이 활성화됐습니다. 눌러주세요.")
            if state["start"] and state["user"]:
                log("\n" + bar)
                log(" 조회 시작")
                log(bar)
                try:
                    rows, failed = run_batch(ctx.request, state["user"], companies,
                                             st, ed, log, notify)
                except KeyboardInterrupt:
                    log("\n중단됨. 지금까지 받은 결과만 저장합니다.")
                    rows, failed = PARTIAL["rows"], PARTIAL["failed"]
                notify("완료: 성공 %d / 실패 %d" % (len(rows), len(failed)))
                break
            try:
                pages[0].wait_for_timeout(400)
            except Exception:
                break
            if time.time() - started > 3600:
                log("1시간이 지나 종료합니다.")
                break

        if rows is not None:
            path = write_excel(rows, failed, st, ed, log)
            log("성공 %d / 실패 %d" % (len(rows), len(failed)))
            for b, m, why in failed:
                log("  실패: %s %s - %s" % (b, m, why))
            notify("저장 완료: %s" % path.name)
            try:
                pages = [x for x in ctx.pages if not x.is_closed()]
                if pages:
                    pages[0].wait_for_timeout(4000)
            except Exception:
                pass
        else:
            log("조회를 실행하지 않고 종료했습니다.")

        try:
            ctx.close()
        except Exception:
            pass

    log.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
