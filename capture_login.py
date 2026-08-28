# -*- coding: utf-8 -*-
"""
1단계: 로그인 + 통신 캡처 + 진단

브라우저를 띄우면 사람이 직접 로그인하고 파생상품 조회를 해본다.
그동안 오간 통신과 브라우저 오류를 전부 파일로 기록한다.

끝낼 때는 콘솔이 아니라 '브라우저 창을 닫으면' 된다.

실행:
    python capture_login.py              # Playwright 전용 Chromium 사용
    python capture_login.py --attach     # 평소 쓰는 Chrome 에 붙는다

--attach 를 쓰려면 Chrome 을 먼저 이렇게 띄운다 (기존 Chrome 은 모두 종료):
    "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe" ^
        --remote-debugging-port=9222 --user-data-dir="C:\\chrome-kfb"
"""
import sys, os, re, json, time, base64, datetime, pathlib, argparse

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from playwright.sync_api import sync_playwright

BASE = pathlib.Path(__file__).parent
PROFILE_DIR = BASE / "browser_profile"
CAPTURE_DIR = BASE / "capture"
LOG_DIR = BASE / "logs"
START_URL = "https://bizinfo.kfb.or.kr/biz/index.html"

POPUP_URL = ("https://bizinfo.kfb.or.kr/biz/popup.html"
             "?formname=BIZKI%3A%3ABIZKI_T0101.xfdl"
             "&framename=KI01_BIZKI_T0101.xfdl")

SKIP_EXT = (".js", ".css", ".png", ".jpg", ".jpeg", ".gif", ".svg",
            ".ico", ".woff", ".woff2", ".ttf", ".map")

# 조회 화면은 URL 을 직접 열 수 없다.
# Nexacro 팝업은 부모 앱이 먼저 팝업 프레임을 등록해야 열리며,
# 등록 없이 popup.html 을 직접 열면 창이 뜨자마자 이렇게 죽는다:
#   Cannot read properties of undefined (reading 'KI01_BIZKI_T0101.xfdl')
#   Cannot read properties of null (reading '_linked_window')
# 따라서 조회 화면은 반드시 사이트 메뉴로 열어야 한다.



# --------------------------------------------------------------------------
# 로그 (콘솔 + 파일 동시 기록. 콘솔 한글이 깨져도 파일은 멀쩡하다)
# --------------------------------------------------------------------------
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


def is_noise(url):
    low = url.split("?")[0].lower()
    if "/nexacro14lib/" in low:
        return True
    return low.endswith(SKIP_EXT)


def body_repr(raw):
    if raw is None:
        return {"text": None, "b64": None, "len": 0}
    out = {"len": len(raw), "text": None, "b64": None}
    for enc in ("utf-8", "euc-kr"):
        try:
            t = raw.decode(enc)
            ctrl = sum(1 for c in t[:2000] if ord(c) < 9 or 11 <= ord(c) < 32)
            if ctrl < max(1, len(t[:2000])) * 0.05:
                out["text"] = t
                out["encoding"] = enc
                return out
        except Exception:
            continue
    out["b64"] = base64.b64encode(raw).decode("ascii")
    out["latin1"] = raw.decode("latin-1", errors="replace")
    return out


def page_tag(resp):
    """이 응답이 어느 창에서 났는지. 팝업 구분용."""
    try:
        u = resp.request.frame.page.url
        return "팝업" if "popup.html" in u else "본창"
    except Exception:
        return "?"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--attach", action="store_true",
                    help="localhost:9222 의 실제 Chrome 에 붙는다")
    ap.add_argument("--port", type=int, default=9222)
    args = ap.parse_args()

    CAPTURE_DIR.mkdir(exist_ok=True)
    LOG_DIR.mkdir(exist_ok=True)
    PROFILE_DIR.mkdir(exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    outfile = CAPTURE_DIR / ("capture_" + stamp + ".jsonl")
    log = Log(LOG_DIR / ("session_" + stamp + ".log"))
    fh = open(outfile, "w", encoding="utf-8")
    n = {"all": 0, "do": 0, "do_popup": 0, "fail": 0}

    def on_response(resp):
        try:
            url = resp.url
            if is_noise(url):
                return
            req = resp.request
            try:
                raw = resp.body()
            except Exception:
                raw = None
            post = req.post_data_buffer
            where = page_tag(resp)
            rec = {
                "ts": datetime.datetime.now().isoformat(timespec="seconds"),
                "where": where,
                "method": req.method, "url": url, "status": resp.status,
                "req_headers": dict(req.headers),
                "req_body": body_repr(post) if post else None,
                "resp_headers": dict(resp.headers),
                "resp_body": body_repr(raw),
            }
            fh.write(json.dumps(rec, ensure_ascii=False) + "\n")
            fh.flush()
            n["all"] += 1
            name = url.split("/")[-1].split("?")[0]
            if url.split("?")[0].endswith(".do"):
                n["do"] += 1
                if where == "팝업":
                    n["do_popup"] += 1
                txt = rec["resp_body"].get("text") or ""
                rows = txt.count("<Row>")
                ec = re.search(r'id="ErrorCode"[^>]*>([^<]*)<', txt)
                em = re.search(r'id="ErrorMsg"[^>]*>([^<]*)<', txt)
                extra = ""
                if ec and ec.group(1).strip() not in ("0", ""):
                    extra = "  !! ErrorCode=%s %s" % (ec.group(1),
                                                      em.group(1) if em else "")
                log("  [업무호출 %3d][%s] %-30s -> %s, %d행%s"
                    % (n["do"], where, name[:30], resp.status, rows, extra))
            else:
                log("  [기타 %3d][%s] %s %s %s"
                    % (n["all"], where, req.method, resp.status, url[:95]))
        except Exception as e:
            log("  (캡처 실패: %s)" % e)

    def on_failed(req):
        n["fail"] += 1
        try:
            log("  [요청실패] %s %s :: %s"
                % (req.method, req.url[:95],
                   req.failure if isinstance(req.failure, str) else req.failure))
        except Exception:
            pass

    def wire_page(pg):
        pg.on("console", lambda m: log("  [브라우저:%s] %s" % (m.type, m.text[:200]))
              if m.type in ("error", "warning") else None)
        pg.on("pageerror", lambda e: log("  [스크립트오류] " + str(e)[:250]))
        pg.on("dialog", lambda d: (log("  [알림창] %s : %s" % (d.type, d.message)),
                                   d.accept()))
        pg.on("close", lambda _: log("  << 창 닫힘"))

    def on_page(pg):
        log("  >> 새 창/탭 열림: " + pg.url[:110])
        wire_page(pg)
        try:
            pg.bring_to_front()
        except Exception:
            pass

    with sync_playwright() as p:
        if args.attach:
            log(" 실제 Chrome (localhost:%d) 에 접속합니다..." % args.port)
            browser = p.chromium.connect_over_cdp("http://localhost:%d" % args.port)
            ctx = browser.contexts[0] if browser.contexts else browser.new_context()
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
        else:
            ctx = p.chromium.launch_persistent_context(
                user_data_dir=str(PROFILE_DIR),
                headless=False, locale="ko-KR", timezone_id="Asia/Seoul",
                viewport=None,
                args=["--disable-blink-features=AutomationControlled",
                      "--disable-popup-blocking", "--start-maximized"])
            page = ctx.pages[0] if ctx.pages else ctx.new_page()

        ctx.on("response", on_response)
        ctx.on("requestfailed", on_failed)
        ctx.on("page", on_page)
        for pg in ctx.pages:
            wire_page(pg)

        if not args.attach:
            page.goto(START_URL, wait_until="domcontentloaded")

        bar = "=" * 78
        log(bar)
        log(" 브라우저 준비 완료.")
        log("")
        log("   1) 로그인 (ID/PW + 추가 인증 전부)")
        log("   2) 파생상품 정보조회 화면을 >> 사이트 메뉴로 << 여세요.")
        log("      URL 을 직접 여는 방법은 동작하지 않습니다.")
        log("      (Nexacro 팝업은 부모 앱이 등록해야만 열립니다)")
        log("   3) 사업자등록번호를 넣고 조회를 누르세요.")
        log("      조회가 안 되더라도 그대로 두고, 아래 로그를 확인하세요.")
        log("   4) 다 됐으면 >> 브라우저 창을 닫으세요 <<")
        log("")
        log(" 로그   : " + str(log.path))
        log(" 캡처   : " + str(outfile))
        log(bar)

        closed = {"v": False}
        try:
            ctx.on("close", lambda _: closed.__setitem__("v", True))
        except Exception:
            pass
        last_save = 0.0
        started = time.time()
        shots = []
        while not closed["v"]:
            try:
                pages = [x for x in ctx.pages if not x.is_closed()]
                if not pages:
                    break
                # 창이 닫히기 전에 화면을 계속 찍어둔다 (마지막 상태 보존용)
                if time.time() - last_save > 10:
                    last_save = time.time()
                    for i, pg in enumerate(pages):
                        try:
                            tag = "popup" if "popup.html" in pg.url else "main"
                            sp = CAPTURE_DIR / ("shot_%s_%d.png" % (tag, i))
                            pg.screenshot(path=str(sp))
                            if str(sp) not in shots:
                                shots.append(str(sp))
                        except Exception:
                            pass
                    if not args.attach:
                        try:
                            ctx.storage_state(path=str(BASE / "session_state.json"))
                        except Exception:
                            pass
                pages[0].wait_for_timeout(400)
            except Exception:
                break
            if time.time() - started > 3600:
                log(" 1시간이 지나 자동 종료합니다.")
                break

        if not args.attach:
            try:
                ctx.storage_state(path=str(BASE / "session_state.json"))
            except Exception:
                pass
            try:
                ctx.close()
            except Exception:
                pass

    fh.close()
    log("")
    log("캡처 %d건 (업무호출 %d건 / 그중 팝업 %d건, 요청실패 %d건)"
        % (n["all"], n["do"], n["do_popup"], n["fail"]))
    log("저장 -> " + str(outfile))
    if shots:
        log("마지막 화면 스냅샷: " + ", ".join(shots))
    summarize(outfile, n, log)
    log.close()
    print("\n>> 이 파일을 알려주세요: %s" % log.path)


def summarize(outfile, n, log):
    bar = "=" * 78
    log("\n" + bar)
    log(" 업무 호출(.do) 요약")
    log(bar)
    seen, order = {}, []
    for line in open(outfile, encoding="utf-8"):
        try:
            r = json.loads(line)
        except Exception:
            continue
        u = r["url"].split("?")[0]
        if u.endswith(".do"):
            if u not in seen:
                seen[u] = []
                order.append(u)
            seen[u].append(r)
    for u in order:
        recs = seen[u]
        log("\n  * %s   (%d회, 창=%s)"
            % (u, len(recs), recs[-1].get("where", "?")))
        r = recs[-1]
        body = (r.get("req_body") or {}).get("text") or ""
        vals = re.findall(r'<Col id="([^"]+)">([^<]*)</Col>', body)
        if vals:
            log("      보낸 값   : "
                + ", ".join("%s=%s" % (k, v) for k, v in vals[:25]))
        rb = (r.get("resp_body") or {}).get("text") or ""
        rcols = re.findall(r'<Column id="([^"]+)"', rb)
        if rcols:
            log("      받은 컬럼 : " + ", ".join(rcols[:30]))
            log("      받은 행수 : %d" % rb.count("<Row>"))

    # ---- 자동 판정 -------------------------------------------------------
    log("\n" + bar)
    log(" 진단")
    log(bar)
    if n["do"] == 0:
        log("  조회 요청(.do)이 한 건도 나가지 않았습니다.")
        log("  -> 화면이 서버에 요청을 보내기 전 단계에서 막힌 것입니다.")
        log("     보통 (a) 필수 입력값 누락으로 화면단 검증에서 걸렸거나,")
        log("          (b) 팝업을 직접 URL로 열어 부모 앱과 연결되지 않았거나,")
        log("          (c) 화면 스크립트 오류입니다. 위 [스크립트오류] 를 보세요.")
    elif n["do_popup"] == 0:
        log("  본창에서는 요청이 나갔지만, 팝업(조회 화면)에서는 0건입니다.")
        log("  -> 조회 화면이 부모 앱과 제대로 연결되지 않았을 가능성이 큽니다.")
        log("     사이트 메뉴를 통해 여는 방법으로 다시 시도해 보세요.")
    else:
        log("  팝업에서 조회 요청이 %d건 나갔습니다." % n["do_popup"])
        log("  -> 통신은 되고 있습니다. 위 '받은 행수'와 ErrorCode 를 확인하세요.")
        log("     행수가 0이면 조건에 맞는 데이터가 없거나 파라미터가 다른 것입니다.")
    log(bar)


if __name__ == "__main__":
    main()
