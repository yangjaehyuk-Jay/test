# -*- coding: utf-8 -*-
"""
Nexacro XML 데이터셋 프로토콜 처리.

은행연합회 통합업무시스템은 화면(Nexacro)과 서버가 아래 형식의 XML을 주고받는다.
이 모듈은 그 XML을 만들고(build_request) 해석하는(parse_response) 일만 한다.
사이트 고유 로직은 들어있지 않다.

    요청  POST /<서비스>.do   Content-Type: text/xml
      <Root xmlns="http://www.nexacroplatform.com/platform/dataset">
        <Parameters><Parameter id="k" type="string">v</Parameter></Parameters>
        <Dataset id="ds_search">
          <ColumnInfo><Column id="BIZNO" type="STRING" size="256"/></ColumnInfo>
          <Rows><Row><Col id="BIZNO">1234567890</Col></Row></Rows>
        </Dataset>
      </Root>

    응답  같은 형식. Parameters 안의 ErrorCode 가 0 이면 정상.
"""
from __future__ import annotations

import xml.etree.ElementTree as ET
from typing import Any, Iterable, Mapping, Sequence

NS = "http://www.nexacroplatform.com/platform/dataset"
_NSMAP = {"n": NS}


# --------------------------------------------------------------------------
# 요청 생성
# --------------------------------------------------------------------------
def _esc(v: Any) -> str:
    """XML 텍스트 이스케이프. Nexacro는 공백을 &#32;로 쓰기도 하지만 필수는 아니다."""
    if v is None:
        return ""
    s = str(v)
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
             .replace('"', "&quot;"))


def build_request(datasets: Mapping[str, Sequence[Mapping[str, Any]]],
                  params: Mapping[str, Any] | None = None,
                  columns: Mapping[str, Sequence[str]] | None = None,
                  col_size: int = 256) -> bytes:
    """
    datasets : {"ds_search": [{"BIZNO": "1234567890"}, ...]}
    params   : {"SVC_ID": "..."} 같은 Parameters 영역
    columns  : 컬럼 순서/집합을 강제하고 싶을 때 {"ds_search": ["BIZNO","YMD"]}.
               생략하면 행 딕셔너리의 키에서 자동으로 뽑는다.
    """
    out = ['<?xml version="1.0" encoding="UTF-8"?>',
           f'<Root xmlns="{NS}">']

    if params:
        out.append("\t<Parameters>")
        for k, v in params.items():
            out.append(f'\t\t<Parameter id="{_esc(k)}" type="string">{_esc(v)}</Parameter>')
        out.append("\t</Parameters>")
    else:
        out.append("\t<Parameters />")

    for ds_id, rows in datasets.items():
        rows = list(rows)
        if columns and ds_id in columns:
            cols = list(columns[ds_id])
        else:
            cols, seen = [], set()
            for r in rows:
                for k in r:
                    if k not in seen:
                        seen.add(k)
                        cols.append(k)

        out.append(f'\t<Dataset id="{_esc(ds_id)}">')
        out.append("\t\t<ColumnInfo>")
        for c in cols:
            out.append(f'\t\t\t<Column id="{_esc(c)}" type="STRING" size="{col_size}" />')
        out.append("\t\t</ColumnInfo>")
        out.append("\t\t<Rows>")
        for r in rows:
            out.append("\t\t\t<Row>")
            for c in cols:
                if c in r and r[c] is not None:
                    out.append(f'\t\t\t\t<Col id="{_esc(c)}">{_esc(r[c])}</Col>')
            out.append("\t\t\t</Row>")
        out.append("\t\t</Rows>")
        out.append("\t</Dataset>")

    out.append("</Root>")
    return "\n".join(out).encode("utf-8")


# --------------------------------------------------------------------------
# 응답 해석
# --------------------------------------------------------------------------
class NexacroError(RuntimeError):
    def __init__(self, code, msg, raw=None):
        super().__init__(f"서버 오류 ErrorCode={code} ErrorMsg={msg}")
        self.code, self.msg, self.raw = code, msg, raw


def _find(el, tag):
    """네임스페이스가 있든 없든 찾는다 (서버가 가끔 ns 없이 보낸다)."""
    r = el.findall(f"n:{tag}", _NSMAP)
    return r if r else el.findall(tag)


def parse_response(data: bytes | str) -> dict:
    """
    반환: {"params": {...}, "datasets": {ds_id: [ {col: val}, ... ]}}
    """
    if isinstance(data, bytes):
        data = data.decode("utf-8", errors="replace")
    root = ET.fromstring(data)

    params: dict[str, str] = {}
    for ps in _find(root, "Parameters"):
        for p in _find(ps, "Parameter"):
            params[p.get("id")] = (p.text or "")

    datasets: dict[str, list[dict]] = {}
    for ds in _find(root, "Dataset"):
        ds_id = ds.get("id")
        cols: list[str] = []
        for ci in _find(ds, "ColumnInfo"):
            for c in _find(ci, "Column"):
                cols.append(c.get("id"))
        rows: list[dict] = []
        for rs in _find(ds, "Rows"):
            for row in _find(rs, "Row"):
                rec = {c: None for c in cols}
                for col in _find(row, "Col"):
                    rec[col.get("id")] = col.text
                rows.append(rec)
        datasets[ds_id] = rows

    return {"params": params, "datasets": datasets}


def check(parsed: dict, raw=None) -> dict:
    """ErrorCode가 0이 아니면 예외를 던지고, 정상이면 그대로 돌려준다."""
    code = parsed["params"].get("ErrorCode")
    if code is not None and str(code).strip() not in ("0", ""):
        raise NexacroError(code, parsed["params"].get("ErrorMsg", ""), raw)
    return parsed


def to_frame(parsed: dict, ds_id: str | None = None):
    """가장 행이 많은 데이터셋(또는 지정한 것)을 pandas DataFrame으로."""
    import pandas as pd
    dss = parsed["datasets"]
    if ds_id is None:
        if not dss:
            return pd.DataFrame()
        ds_id = max(dss, key=lambda k: len(dss[k]))
    return pd.DataFrame(dss.get(ds_id, []))
