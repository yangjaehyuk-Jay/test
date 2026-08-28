# 은행연합회 파생상품 정보 일일 조회

`bizinfo.kfb.or.kr` (은행연합회 통합업무시스템)에서 약 60개 기업의
파생거래 정보를 매일 조회해 엑셀로 저장한다.

## 사용법

```
cd "C:\Users\infomax\Documents\클로드\KFB_파생거래"
python scrape.py
```

1. 브라우저가 열리면 **로그인** (ID/PW + 추가 인증)
2. 로그인이 끝나면 우측 하단 **`▶ 자동 조회 시작`** 버튼이 파랗게 켜진다
3. 버튼을 누르면 `companies.csv` 의 기업을 차례로 조회한다
4. 끝나면 `output/파생거래_YYYYMMDD.xlsx` 저장

옵션:

```
python scrape.py --from 20260101 --to 20260828   # 조회기간 (기본: 올해 1월 1일 ~ 오늘)
python scrape.py --delay 4                        # 요청 간격 초 (기본 2.5)
```

조회 대상은 `companies.csv` 에 한 줄에 하나씩. 하이픈은 있어도 없어도 된다.

```
101-86-66838,삼양사
4548703296,삼양바이오팜
```

## 알아낸 API

캡처(`capture_login.py`)로 확인한 실제 통신 구조.

### 조회

```
POST https://bizinfo.kfb.or.kr/bizkiR0101SelectList.do
Content-Type: text/xml
```

요청 `ds_search` (1행):

| 컬럼 | 값 |
|---|---|
| `USER_ID` / `USER_NM` | 로그인 사용자 |
| `USER_ORG_CD` / `USER_ORG_NM` | 소속 기관 |
| `KIKO_ID_NO_TP_CD` | `2` = 사업자등록번호 (`1` = 주민등록번호) |
| `KIKO_ID_NO` | **`"999"` + 사업자등록번호 10자리** (13자리 조회요청 식별번호) |
| `INQ_DV_CD` | `1` |
| `INQ_ST_DT` / `INQ_ED_DT` | 조회기간 `YYYYMMDD` |
| `IS_IN_USER` | `N` |

`USER_*` 네 개는 로그인 응답 `bizliS0102Login2.do` 의 `ds_list` 에서 그대로 나온다.
스크립트가 로그인 응답을 엿들어 자동으로 채운다.

응답 `ds_etc` (조회 상태):

| 컬럼 | 뜻 |
|---|---|
| `RSP_CD` / `RSP_NM` | `000` 정상 / `401` 조회 해당자료 없음 |
| `KIKO_CUST_NM` | 기업명 (전각문자로 옴 → NFKC 정규화함) |
| `KIKO_ID_NO` | 사업자번호 (하이픈 포함) |
| `RETRIEVE_DTTM` | 조회 시각 |
| `BNS_AMT_USD_CVT_AMT` | USD 환산금액 |

응답 `ds_list` (실제 잔액. 접두어 = 상품군):

- `FX_*` 선물환 &nbsp; `OP_*` 옵션 &nbsp; `CRS_*` 통화스왑
- `_VLD_CNT` 유효건수, `_BUY_BAL` / `_SELL_BAL` 매수·매도 잔액,
  `_TRN_LST_UPD_DT` 최종갱신일
- 옵션만 세분: `OP_BUY_CALL_BAL`, `OP_BUY_PUT_BAL`,
  `OP_SELL_CALL_BAL`, `OP_SELL_PUT_BAL`

### 감사로그

```
POST https://bizinfo.kfb.or.kr/bizcmEventLoggerInsert.do
ds_log: USER_ID, USER_NM, SCREEN_ID=/BIZKI/BIZKI_R0101,
        SCREEN_NM=파생상품(TAB_01), ACT_NM=화면오픈|내역조회
```

사이트가 조회 이력을 남기는 호출이다. **자동 조회도 이걸 똑같이 보낸다.**
화면으로 조회할 때와 감사 흔적을 동일하게 유지하기 위해서다.

## 왜 이런 구조인가

| 항목 | 내용 |
|---|---|
| 화면 기술 | Nexacro 14 HTML5. 데이터는 평문 XML 데이터셋 |
| 조회 화면 | **별도 팝업 창**. 반드시 사이트 메뉴로만 열린다 (아래 참고) |
| 로그인 | ID/PW + 추가 인증 → 무인 자동화 불가. 사람이 1회 로그인 |
| 세션 | 화면에 "20분 후 자동 로그아웃" 표시 |
| 서버 | **연속 요청을 차단한다.** 막히면 XML 대신 236바이트 안내 HTML이 온다 |
| 필수 프로그램 | nProtect Online Security, MarkAny e-PageSafer (이 PC 설치완료) |

### 조회 화면은 URL 로 직접 열 수 없다

`popup.html?formname=BIZKI::BIZKI_T0101.xfdl&framename=KI01_BIZKI_T0101.xfdl` 을
브라우저에 그냥 붙여넣으면 창은 뜨지만 곧바로 죽는다.

```
Cannot read properties of undefined (reading 'KI01_BIZKI_T0101.xfdl')
Cannot read properties of null (reading '_linked_window')
```

Nexacro 팝업은 부모 앱이 먼저 팝업 프레임을 등록해야 열린다.
URL 만 여는 방식은 그 등록을 건너뛰므로 프레임과 부모 연결을 못 찾는다.
**조회 화면이 필요하면 사이트 메뉴로 열어야 한다.**

다만 `scrape.py` 는 이 화면을 열지 않는다. 화면이 보내는 XML 요청을
직접 보내기 때문에 팝업이 떠 있든 말든 상관없다.

화면을 클릭해서 긁지 않는다. Nexacro 그리드는 DOM이 수시로 바뀌어 취약하다.
대신 화면이 쓰는 XML 요청을 그대로 보낸다.

요청 간 기본 2.5초를 쉬고, 실패하면 8초 → 16초로 물러서며 3회까지 재시도한다.
56개 기업이면 감사로그 호출까지 합쳐 대략 5~6분 걸린다.

## 파일

| 파일 | 역할 |
|---|---|
| `scrape.py` | **본 프로그램.** 로그인 대기 → 일괄 조회 → 엑셀 저장 |
| `companies.csv` | 조회 대상 사업자등록번호 목록 |
| `nexacro.py` | Nexacro XML 요청 생성 / 응답 파싱 |
| `capture_login.py` | 통신 캡처·진단 도구. API가 바뀌었을 때 다시 뜯어볼 용도 |
| `output/` | 결과 엑셀 (`파생거래_YYYYMMDD.xlsx`) |
| `logs/` | 실행 로그 |
| `browser_profile/` | 로그인 세션. **외부 공유 금지** |

엑셀은 3개 시트다. `파생거래`(결과), `조회정보`(조회일시·기간·성공/실패 수),
`실패`(실패한 기업과 사유. 실패가 있을 때만).

## 문제가 생기면

먼저 `logs/scrape_*.log` 를 본다.

| 증상 | 원인 |
|---|---|
| `XML 아님 (차단 추정)` | 서버가 막았다. `--delay 5` 로 늘린다 |
| 버튼이 안 켜짐 | 로그인 응답을 못 봤다. 로그아웃 후 다시 로그인 |
| `결과코드 401` | 정상 동작. 그 기업은 해당 자료가 없다 |
| 대량 실패 | 세션 만료 가능성. 다시 실행해 재로그인 |

중간에 끊겨도 `output/_진행중_YYYYMMDD.csv` 에 조회한 만큼 이미 저장돼 있다.

API 자체가 바뀐 것 같으면 `python capture_login.py` 로 다시 캡처해서
`.do` 호출 요약을 비교한다.

## 남은 작업

- [x] `companies.csv` 56개 등록 완료 (형식오류 0, 중복 0, CP949 인코딩)
- [ ] 첫 실전 실행으로 56개 전량 확인
- [ ] (선택) 전일 대비 변동분 비교 시트
- [ ] (선택) Windows 작업 스케줄러 등록 — 단, 로그인이 수동이라
      "정해진 시각에 창을 띄워주는" 용도까지만 가능하다
