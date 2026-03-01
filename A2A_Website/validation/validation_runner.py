import os, json
import pandas as pd
import sys
import re, shutil, pkgutil
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding='utf-8')

from functools import lru_cache

@lru_cache(maxsize=64)
def _read_excel_cached(path: str, sheet_name: str, header: int, usecols_key: tuple | None):
    """
    엑셀 시트 읽기를 LRU 캐시로 가속.
    - path, sheet_name, header, usecols 튜플을 키로 캐싱
    - dtype=str로 통일하여 타입 추론 비용 최소화
    """
    usecols = list(usecols_key) if usecols_key else None
    df = pd.read_excel(path, sheet_name=sheet_name, header=header, dtype=str, usecols=usecols)
    return df

def _bundle_path(rel: str) -> str:
    """PyInstaller 번들 내부(읽기 전용) 상대 경로 → 절대경로"""
    return os.path.join(getattr(sys, "_MEIPASS", ""), rel)

def _external_path(rel: str) -> str:
    """exe 가 놓인 폴더 기준(쓰기 가능) 상대 경로 → 절대경로"""
    base = sys.executable if getattr(sys, "frozen", False) else __file__
    return os.path.abspath(os.path.join(os.path.dirname(base), rel))

# 1) read-only 템플릿  / 2) run-time read-write 디렉터리
TEMPLATE_DIR = _bundle_path("configured_steps")      # exe 에 묶여있는 초기 스텝
DATA_DIR     = _external_path("configured_steps")    # 실제 편집·저장용 스텝

def _ensure_data_dir():
    """처음 실행 시 TEMPLATE_DIR → DATA_DIR 로 복사(이미 있으면 skip)"""
    if not os.path.isdir(DATA_DIR):
        shutil.copytree(TEMPLATE_DIR, DATA_DIR, dirs_exist_ok=True)
        print(f"[INIT] configured_steps 템플릿을 초기화했습니다 → {DATA_DIR}")

_ensure_data_dir()
# ==============================================

# ✅ 이 파일에서 사용하던 BASE_DIR 을 모두 DATA_DIR 로 대체
BASE_DIR = DATA_DIR
print("### validation_runner LOADED FROM:", __file__)
print("### BASE_DIR:", BASE_DIR)

def hawb_clean_key(s: pd.Series, *, preserve_nec_hyphen: bool = True) -> pd.Series:
    """
    HAWB 매칭용 정규화:
      - 앞뒤 공백 제거, 대문자화, 내부 공백 제거
      - 기본: 모든 '-' 제거
      - preserve_nec_hyphen=True면 'NEC-' 접두 하이픈만 보존
        예) 'NEC- 123 45-67' -> 'NEC-1234567'
            'SEL-1003 0259'  -> 'SEL10030259'
    """
    s = s.astype("string").str.normalize("NFKC").str.strip().str.upper()
    s_nospace = s.str.replace(r"\s+", "", regex=True)  # 모든 공백 제거

    if not preserve_nec_hyphen:
        return s_nospace.str.replace("-", "", regex=False)

    # 공백 제거된 상태에서 NEC- 접두 여부 판단
    nec_mask = s_nospace.str.startswith("NEC-")

    # 일단 모든 '-' 제거
    no_dash = s_nospace.str.replace("-", "", regex=False)

    # NEC- 인 경우에만 'NEC-' 복원: 'NEC' + '-' + 나머지
    # (no_dash는 하이픈이 제거됐으므로 인덱스 3 이후 붙여주면 됨)
    return no_dash.where(~nec_mask, "NEC-" + no_dash.str.slice(3))

# ──────────────────────────────────────────────
# 1.  Step-별 로직 함수들
# ──────────────────────────────────────────────
def create_columns(df: pd.DataFrame, cfg):
    """지정된 컬럼이 없으면 만들어 둠 (문자열·리스트 모두 허용)"""
    cols = cfg.get("columns", [])
    if isinstance(cols, str):          # ★ 한-개만 전달된 경우
        cols = [cols]
    for col in cols:
        if col not in df.columns:
            df[col] = None
    return df


def sort_hawb(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    """
    Generic sort step.

    cfg accepted shapes:
      { "column": "HAWB / Bill of Lading #", "ascending": true }
      { "column": "HAWB / Bill of Lading #", "order": "desc" }
      { "columns": ["HAWB / Bill of Lading #", "Custom Form Declaration Date"],
        "ascending": [true, false] }
      { "columns": [
            {"name": "HAWB / Bill of Lading #", "ascending": true},
            {"name": "Custom Form Declaration Date", "ascending": false}
        ] }
      { "numeric": true }   # attempt numeric conversion on first column
    """
    if cfg is None:
        cfg = {}

    # ---- 1) Normalize column list -----------------------------------------
    cols = []
    asc_list = []

    if "columns" in cfg:
        c = cfg["columns"]
        if isinstance(c, list) and c and isinstance(c[0], dict):
            # list of dicts
            for coldef in c:
                cols.append(coldef["name"])
                asc_list.append(bool(coldef.get("ascending", True)))
        else:
            # list of column names
            cols = list(c)
            asc_raw = cfg.get("ascending", True)
            if isinstance(asc_raw, list):
                asc_list = [bool(a) for a in asc_raw]
            else:
                asc_list = [bool(asc_raw)] * len(cols)
    elif "column" in cfg:
        cols = [cfg["column"]]
        # order or ascending flags
        if "order" in cfg:
            asc_list = [str(cfg["order"]).lower() != "desc"]
        else:
            asc_list = [bool(cfg.get("ascending", True))]
    else:
        # fallback default
        cols = ["HAWB"]
        asc_list = [True]

    # 길이 맞추기
    if len(asc_list) != len(cols):
        asc_list = asc_list + [asc_list[-1]] * (len(cols) - len(asc_list))

    # ---- 2) Defensive: missing columns → warn & skip missing ones ----------
    missing = [c for c in cols if c not in df.columns]
    if missing:
        print(f"[sort_hawb] WARNING: missing columns {missing} – ignoring them.")
        cols = [c for c in cols if c in df.columns]
        asc_list = [a for c, a in zip(cols, asc_list) if c in df.columns]

    if not cols:
        print("[sort_hawb] No valid columns to sort by. Returning unchanged.")
        return df

    # ---- 3) Optional numeric coercion on *first* column --------------------
    if cfg.get("numeric"):
        first = cols[0]
        try:
            df[first + "__num"] = pd.to_numeric(df[first], errors="coerce")
            # sort using numeric col first, then rest
            df = df.sort_values(
                by=[first + "__num"] + cols[1:],
                ascending=asc_list, kind="mergesort"  # stable
            ).drop(columns=[first + "__num"])
            return df.reset_index(drop=True)
        except Exception as e:
            print(f"[sort_hawb] numeric coercion failed: {e}; falling back to text.")

    # ---- 4) Regular sort ---------------------------------------------------
    df = df.sort_values(by=cols, ascending=asc_list, kind="mergesort")
    return df.reset_index(drop=True)


import pandas as pd
import re

def _hawb_key(s: pd.Series, normalize: bool = True) -> pd.Series:
    """HAWB 비교용 키 생성: 조인과 동일 규칙으로!"""
    s = s.astype("string")
    if not normalize:
        return s
    # 유니코드 정규화 + 양끝 공백 제거 + 대문자화
    s = (s.str.normalize("NFKC")
           .str.strip()
           .str.upper())
    # 다양한 대시(‐-‒–—―− 포함) 제거
    s = s.str.replace(r"[\u2010-\u2015\u2212-]", "", regex=True)
    # 내부 공백 제거(필요시 해제 가능)
    s = s.str.replace(r"\s+", "", regex=True)
    return s

def unique_hawb(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    col       = cfg.get("column", "HAWB")
    group_na  = bool(cfg.get("group_na", True))

    if col not in df.columns:
        raise KeyError(f"unique_hawb: '{col}' column not found")

    out = df.copy()
    # ⬇ 기존 수기 정규화 대신, NEC- 보존 규칙 포함한 헬퍼 사용
    out["__HAWB_KEY__"] = hawb_clean_key(out[col], preserve_nec_hyphen=True)

    out = out.sort_values(["__HAWB_KEY__"], kind="mergesort").reset_index(drop=True)
    grp = out["__HAWB_KEY__"].fillna("__NA__") if group_na else out["__HAWB_KEY__"]
    first_mask = ~grp.duplicated(keep="first")

    # Unique Order 로직은 기존 그대로 유지
    out["Unique Order"] = first_mask.astype(int)
    # Total Order 는 모든 행에서 1로 고정
    out["Total Order"]  = 1

    return out.drop(columns="__HAWB_KEY__")




def remove_by_remarks(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    col   = cfg.get("column", "Remarks")
    raw   = cfg["values"]

    # 문자열이면 쉼표 기준 split
    if isinstance(raw, str):
        raw_vals = [v.strip() for v in raw.split(",")]
    else:
        raw_vals = raw

    drop_vals = {v.upper() for v in raw_vals}

    hawb_col = cfg.get("hawb_column", "HAWB / Bill of Lading #")
    mode     = cfg.get("cascade", "row")   # row | hawb

    mask = df[col].astype(str).str.upper().str.strip().isin(drop_vals)

    if mode == "hawb":
        bad_hawb = df.loc[mask, hawb_col].astype(str).str.strip().unique()
        df = df[~df[hawb_col].isin(bad_hawb)].copy()
    else:
        df = df[~mask].copy()

    return df


def wd_filtering(df: pd.DataFrame, cfg):
    """WD 키워드 행 제거"""
    col   = cfg["column"]
    vals  = set(str(v).upper() for v in cfg["values"])
    series = df[col].astype(str).str.upper().str.strip()
    mask   = series.isin(vals)           # 포함된 행
    return df[~mask].copy()              # 제거 후 반환

# ──────────────────────────────────────────────────────────────
# [MONKEY-PATCH] pd.Series.str.contains → 특수문자 자동 이스케이프
# ──────────────────────────────────────────────────────────────
_original_contains = pd.Series.str.contains  # ← 원본 저장

def _literal_contains(self, pat, *args, **kwargs):
    """
    regex=True(기본값)일 때 패턴에 정규식 메타문자가 있으면
    자동으로 re.escape(pat) 처리해서 '글자 그대로' 비교하도록 한다.
    """
    # regex=False 로 직접 호출한 경우는 사용자의 의도이므로 그대로 둔다.
    if kwargs.get("regex", True) and isinstance(pat, str):
        pat = re.escape(pat)
    return _original_contains(self, pat, *args, **kwargs)

# 실제 패치 적용
pd.Series.str.contains = _literal_contains
# ──────────────────────────────────────────────────────────────


def remarks_update(df: pd.DataFrame, cfg):
    """
    df[<target_column>] 컬럼을 룰 기반으로 업데이트.
    - 기본은 Remarks 컬럼
    - 새 포맷: {"rules":[{ "target_column": "Carrier / LSP", … }]}
    - 레거시: {"actions":[…]} (Remarks 전용)
    """
    if "Remarks" not in df.columns:
        df["Remarks"] = None

    # ── (NEW) append 모드 옵션: 기본 ON, 구분자 기본 " /" ─────────
    append_mode = bool(cfg.get("append_mode", True))   # 기본 True
    sep = str(cfg.get("append_sep", " / "))             # 기본 " /"
    sep_esc = re.escape(sep)                           # 중복 체크용

    # ── 0) 스텝 수준 기본 대상 컬럼(없으면 Remarks) ──────────────
    default_target_col = cfg.get("target_column", "Remarks")
    if default_target_col not in df.columns:
        df[default_target_col] = None

    # ── 1) 신형 포맷 ────────────────────────────────────────────
    if "rules" in cfg:
        for rule in cfg["rules"]:
            target = rule["target_value"]

            # 1-A) 룰별 대상 컬럼(없으면 스텝 기본)
            tgt_col = rule.get("target_column", default_target_col)
            if tgt_col not in df.columns:
                df[tgt_col] = None

            mask = pd.Series(True, index=df.index)

            for cond in rule["conditions"]:
                col  = cond["column"]
                op   = cond["operator"]
                vals = [str(v).upper() for v in cond.get("values", [])]

                raw   = df[col]
                upper = raw.astype(str).str.strip().str.upper()

                if   op == "equals":          cond_mask = upper == vals[0]
                elif op in ("in", "IN"):      cond_mask = upper.isin(vals)
                elif op == "not_in":          cond_mask = ~upper.isin(vals)
                elif op == "is_blank":        cond_mask = upper.eq("") | raw.isna()
                elif op == "is_not_blank":    cond_mask = ~(upper.eq("") | raw.isna())
                elif op == "contains":        cond_mask = upper.str.contains(vals[0], na=False)
                elif op == "not_contains":    cond_mask = ~upper.str.contains(vals[0], na=False)
                elif op == "starts_with":     cond_mask = upper.str.startswith(vals[0], na=False)
                elif op == "not_starts_with": cond_mask = ~upper.str.startswith(vals[0], na=False)
                else:
                    raise ValueError(f"Unsupported op: {op}")

                mask &= cond_mask

            # ── 대입 지점 (append 모드 반영)
            if append_mode:
                cur = df[tgt_col].astype("string")
                to_set_empty = mask & (cur.isna() | cur.str.strip().eq(""))

                # NOTE: 전역 monkey-patch로 str.contains가 리터럴화되므로
                #       경계 검사를 위해 컴파일된 정규식을 쓴다(escape 방지).
                pattern = re.compile(rf'(^|{sep_esc}){re.escape(str(target))}({sep_esc}|$)')
                has_token = cur.fillna("").str.contains(pattern, na=False, regex=True)

                to_append = mask & ~to_set_empty & ~has_token

                df.loc[to_set_empty, tgt_col] = str(target)
                df.loc[to_append,    tgt_col] = cur[to_append] + sep + str(target)
            else:
                # 기존 동작: 마지막 규칙이 덮어씀
                df.loc[mask, tgt_col] = target

    # ── 2) 레거시 포맷 ─────────────────────────────────────────
    elif "actions" in cfg:
        for act in cfg["actions"]:
            cond    = act["condition"]
            col     = cond["column"]
            op      = cond["operator"]
            values  = [str(v).upper() for v in cond["values"]]
            target  = act["value"]

            series  = df[col].astype(str).str.strip().str.upper()
            upper   = series  # ← 필수

            if   op == "equals":        mask = upper == values[0]
            elif op == "in":            mask = upper.isin(values)
            elif op == "not in":        mask = ~upper.isin(values)
            elif op == "contains":      mask = upper.str.contains(values[0], na=False)
            elif op == "not equals":    mask = upper != values[0]
            else:
                raise ValueError(f"Unsupported op: {op}")

            if append_mode:
                tgt_col = "Remarks"
                cur = df[tgt_col].astype("string")
                to_set_empty = mask & (cur.isna() | cur.str.strip().eq(""))

                pattern = re.compile(rf'(^|{sep_esc}){re.escape(str(target))}({sep_esc}|$)')
                has_token = cur.fillna("").str.contains(pattern, na=False, regex=True)

                to_append = mask & ~to_set_empty & ~has_token

                df.loc[to_set_empty, tgt_col] = str(target)
                df.loc[to_append,    tgt_col] = cur[to_append] + sep + str(target)
            else:
                df.loc[mask, "Remarks"] = target
    else:
        raise KeyError("remarks_update: neither 'rules' nor 'actions' found in cfg")

    return df


def pbi_multi_merge(
        df: pd.DataFrame,
        cfg: dict,
        pbi_file_path: str
) -> pd.DataFrame:
    """
    여러 PBI 시트를 병합하여 PBI·FG·WD 컬럼을 생성한다.
      • 좌/우 키 정규화 후 조인 (공백 제거, 대문자화, '-' 제거하되 'NEC-' 접두 하이픈은 보존)
      • 접미사(_FG/_WD) 없이 바로 대상 이름(PBI/FG/WD) 부여
      • rename_map 항목이 있으면 최종 수동 오버라이드
      • 성능개선: usecols 최소화 + left_key 사전계산 + map 기반 조인 + 타이밍 로그
    """
    import time

    # ───────────── 공통 유틸 ─────────────
    def _norm(s: pd.Series, *, rm_dash=True):
        # (필터 비교 등 일반 용도) 공백제거 + 대문자화 + 선택적 하이픈 제거
        s = s.astype("string").str.strip().str.upper()
        return s.str.replace("-", "", regex=False) if rm_dash else s

    def hawb_clean_key(s: pd.Series, *, preserve_nec_hyphen: bool = True) -> pd.Series:
        """
        HAWB 매칭용 정규화:
          - 앞뒤 공백 제거, 대문자화
          - 모든 내부 공백 제거
          - 기본: 모든 '-' 제거
          - preserve_nec_hyphen=True면 'NEC-' 접두 하이픈만 보존
            예) 'NEC- 123 45-67' -> 'NEC-1234567'
                'SEL-1003 0259'  -> 'SEL10030259'
        """
        s = s.astype("string").str.normalize("NFKC").str.strip().str.upper()
        s_nospace = s.str.replace(r"\s+", "", regex=True)  # 모든 공백 제거

        if not preserve_nec_hyphen:
            return s_nospace.str.replace("-", "", regex=False)

        nec_mask = s_nospace.str.startswith("NEC-")
        no_dash = s_nospace.str.replace("-", "", regex=False)
        # NEC- 인 경우만 'NEC-' 복원: 'NEC' + '-' + 나머지
        return no_dash.where(~nec_mask, "NEC-" + no_dash.str.slice(3))

    debug = bool(cfg.get("debug_merge", False))
    t0_all = time.perf_counter()

    df = df.copy()

    # 메인 HAWB(과거 호환용 디버그 키)
    main_hawb_col = cfg.get("main_hawb_column", "HAWB")
    if main_hawb_col not in df.columns:
        raise KeyError(f"pbi_multi_merge: '{main_hawb_col}' column not found in DataFrame.")

    # 과거 로그 호환용
    df["__JOIN_KEY"] = _norm(df[main_hawb_col])
    if debug:
        left_na = int(df["__JOIN_KEY"].isna().sum())
        left_empty = int((df["__JOIN_KEY"] == "").sum())
        print(f"[DEBUG] LEFT keys(main='{main_hawb_col}'): total={len(df)}, NA={left_na}, empty='{left_empty}'")

    # ① 좌측 키 사전 계산(소스마다 다른 left_key 대비) — NEC- 접두 하이픈 보존 규칙 적용
    left_keys_cache: dict[str, pd.Series] = {}
    for src in cfg["sources"]:
        lkey = src.get("left_key", main_hawb_col)
        if lkey not in df.columns:
            raise KeyError(f"pbi_multi_merge: left_key '{lkey}' not found in DataFrame columns.")
        if lkey not in left_keys_cache:
            left_keys_cache[lkey] = hawb_clean_key(df[lkey], preserve_nec_hyphen=True)

    # ② 시트 캐시(같은 시트를 여러 번 쓸 때 I/O 절약)
    sheet_cache: dict[str, pd.DataFrame] = {}

    # ──────── source 시트 순회 병합 ────────
    for src in cfg["sources"]:
        t_sheet = time.perf_counter()

        sheet   = src["sheet"]
        l_key   = src.get("left_key", main_hawb_col)
        r_key   = src["right_key"]
        alias   = src.get("alias", sheet.replace(" ", "_").upper())   # NFG / FG / WD …

        keep_raw = src.get("cols", [])
        keep = [keep_raw] if isinstance(keep_raw, str) else list(keep_raw)
        keep = [c for c in keep if c]
        if not keep:
            if debug:
                print(f"[DEBUG] Skip sheet='{sheet}' (no 'cols')")
            continue

        # ── usecols 최소화 구성
        usecols = set([r_key] + keep)
        if src.get("filter_col"):
            usecols.add(src["filter_col"])
        usecols = list(usecols)

        # ① 시트 로드(캐시)  ── ★★ 순서 의존성 해결: 부족 열이 있으면 '합집합'으로 재로딩 ★★
        header = 2  # 기존과 동일(엑셀 3행이 헤더)
        if sheet in sheet_cache:
            pbi = sheet_cache[sheet]
            missing = [c for c in usecols if c not in pbi.columns]
            if missing:
                # 기존 컬럼 ∪ 필요한 컬럼으로 다시 읽어서 캐시 갱신
                union_cols = list(dict.fromkeys(list(pbi.columns) + missing))
                try:
                    pbi = _read_excel_cached(pbi_file_path, sheet, header, tuple(union_cols)).copy()
                except NameError:
                    # LRU 유틸이 없는 환경이면 일반 read_excel 사용
                    try:
                        pbi = pd.read_excel(pbi_file_path, sheet_name=sheet, header=header,
                                            dtype=str, usecols=union_cols)
                    except Exception:
                        # usecols 이름 불일치 등으로 실패하면 최후의 수단: 전체 열 로드
                        pbi = pd.read_excel(pbi_file_path, sheet_name=sheet, header=header, dtype=str)
                except Exception:
                    # usecols 이름 불일치 등으로 실패하면 최후의 수단: 전체 열 로드
                    pbi = pd.read_excel(pbi_file_path, sheet_name=sheet, header=header, dtype=str)
                sheet_cache[sheet] = pbi
                if debug:
                    print(f"[DEBUG] reload '{sheet}' with union usecols ({len(union_cols)}) to cover missing {missing}")
        else:
            t_read = time.perf_counter()
            try:
                pbi = _read_excel_cached(pbi_file_path, sheet, header, tuple(usecols)).copy()  # LRU 캐시 사용
            except NameError:
                pbi = pd.read_excel(pbi_file_path, sheet_name=sheet, header=header,
                                    dtype=str, usecols=usecols)
            except Exception:
                # 첫 로딩도 usecols 문제로 실패하면 전체 열 로드
                pbi = pd.read_excel(pbi_file_path, sheet_name=sheet, header=header, dtype=str)
            sheet_cache[sheet] = pbi

            if debug:
                print(f"[DEBUG] read_excel('{sheet}', usecols={len(usecols)}) : {time.perf_counter() - t_read:.3f}s")

        # ② (선택) 필터 – literal contains/equals/startswith
        if src.get("filter_col"):
            col = src["filter_col"]
            op  = src.get("filter_op", "equals")
            raw_val = str(src.get("filter_val", ""))

            # 컬럼 값: 앞뒤 공백 제거 + 대문자화 + "모든 공백 제거" (하이픈은 그대로 유지)
            col_s = (
                pbi[col].astype("string")
                .str.normalize("NFKC")
                .str.strip()
                .str.upper()
                .str.replace(r"\s+", "", regex=True)  # ← 내부 공백 제거 (예: "8831 3239 6788" → "883132396788")
            )

            # 비교 값도 동일 규칙으로 정규화(공백 제거 + 대문자화)
            val = re.sub(r"\s+", "", raw_val).upper()

            if   op == "equals":
                pbi_f = pbi[col_s == val]
            elif op == "startswith":
                pbi_f = pbi[col_s.str.startswith(val, na=False)]
            elif op == "contains":
                # 리터럴 포함(정규식 아님)
                pbi_f = pbi[col_s.str.contains(val, na=False, regex=False)]
            else:
                pbi_f = pbi
        else:
            pbi_f = pbi

        # ③ 우측 키 생성 — NEC- 접두 하이픈 보존 규칙 적용
        rk = hawb_clean_key(pbi_f[r_key], preserve_nec_hyphen=True)
        pbi_keep = pbi_f.assign(__JOIN_KEY_R=rk)[keep + ["__JOIN_KEY_R"]]

        # 🔒 빈 키 제거
        pbi_keep = pbi_keep[pbi_keep["__JOIN_KEY_R"].notna() & (pbi_keep["__JOIN_KEY_R"] != "")]

        # ④ 타겟 컬럼 이름 정하기(단일/다중 모두 지원)
        tgt_name = alias
        rename_map = {keep[0]: tgt_name} if len(keep) == 1 else {
            c: f"{tgt_name}_{i}" for i, c in enumerate(keep, 1)
        }
        pbi_keep = pbi_keep.rename(columns=rename_map)

        # ⑤ m:1 매핑으로 병합(merge 대신 map; 컬럼별 1:1 매핑 생성)
        #    - 같은 키의 여러 행이 있으면 "해당 컬럼의 첫 번째 비결측값"을 사용
        left_key_series = left_keys_cache[l_key]
        for src_col, dst_col in rename_map.items():
            col_series = pbi_keep[["__JOIN_KEY_R", dst_col]]

            nonnull = col_series[col_series[dst_col].notna() & (col_series[dst_col].astype(str).str.strip() != "")]
            if nonnull.empty:
                continue

            # 키 기준 최초 1건만(빠름) — 정렬 없이 입력 순서 첫 값 보존
            first_per_key = nonnull.drop_duplicates(subset="__JOIN_KEY_R", keep="first")

            # dict 매핑 → 매우 빠른 벡터화 치환
            mapper = pd.Series(first_per_key[dst_col].values, index=first_per_key["__JOIN_KEY_R"].values)

            # 실제 대입 (기존과 동일하게 '새 컬럼'을 생성/갱신)
            df[dst_col] = left_key_series.map(mapper).astype("string")

        if debug:
            print(f"[DEBUG] mapped '{sheet}' (l='{l_key}', r='{r_key}', cols={len(keep)}) : {time.perf_counter() - t_sheet:.3f}s")

    # ─────── post_rules (선택) ───────
    for rule in cfg.get("post_rules", []):
        if not rule.get("when"):
            continue
        mask = df.eval(rule["when"])
        for col, val in rule.get("set", {}).items():
            df.loc[mask, col] = val

    # ─────── rename_map (수동 오버라이드) ───────
    raw_map = cfg.get("rename_map", {})
    custom_map = {item["old"]: item["new"] for item in raw_map} if isinstance(raw_map, list) else raw_map
    if custom_map:
        # ★ 변경: 기존 rename 대신 '업데이트-후-정리' 전략
        for old, new in custom_map.items():
            if old not in df.columns:
                continue
            # new 컬럼이 이미 raw에 있으면: old의 비공백/비결측 값만 new에 덮어쓰기(매칭된 행만 업데이트)
            if new in df.columns:
                s_old = df[old].astype("string")
                mask = s_old.notna() & s_old.str.strip().ne("")
                if mask.any():
                    df.loc[mask, new] = s_old[mask]
                # 보조(old) 컬럼은 정리
                df.drop(columns=[old], inplace=True)
                if debug:
                    print(f"[DEBUG] updated existing '{new}' from '{old}' and dropped '{old}'")
            else:
                # new 컬럼이 없으면 기존처럼 rename
                df.rename(columns={old: new}, inplace=True)
                if debug:
                    print(f"[DEBUG] renamed '{old}' -> '{new}'")

    if debug:
        print(f"[DEBUG] pbi_multi_merge total : {time.perf_counter() - t0_all:.3f}s")

    # 정리
    return df.drop(columns="__JOIN_KEY", errors="ignore")





def wd_hawb_exclude(
    df: pd.DataFrame,
    cfg: dict,
    hawb_excel_path: str
) -> pd.DataFrame:
    """
    WD Power BI 시트에 존재하고, Commodity 열이 'WD' 인 HAWB 전부를
    원본 DataFrame에서 제거한다.
    cfg 예시:
      {
        "wd_sheet": "WD_Hub",
        "wd_column_hawb": "FG_WD_CAPITAL[ORDER_RELEASE_XID]",
        "wd_column_commodity": "FG_WD_CAPITAL[COMMODITY]",
        "wd_commodity_value": "WD",               # (선택) 기본값 WD
        "main_hawb_column": "HAWB / Bill of Lading #"
      }
    """
    wd_sheet   = cfg["wd_sheet"]
    col_hawb   = cfg["wd_column_hawb"]
    col_com    = cfg["wd_column_commodity"]
    header_row = int(cfg.get("header_row", 3))   # ← 기본 2 (0‑기준)
    com_value  = str(cfg.get("wd_commodity_value", "WD")).upper()
    col_raw    = cfg.get("main_hawb_column", "HAWB / Bill of Lading #")

    # ── WD 시트 로드 & 필터 ───────────────────────────────
    wd_df = pd.read_excel(hawb_excel_path, sheet_name=wd_sheet, skiprows=2)
    wd_df.columns = wd_df.columns.str.strip()  # 공백 제거
    

    wd_keys = (
        wd_df.loc[
            wd_df[col_com].astype(str).str.strip().str.upper() == com_value,
            col_hawb
        ]
        .dropna()
        .astype(str).str.strip().str.upper()
        .str.replace("-", "", regex=False)      # 포맷 불일치 대비
        .unique()
    )

    print(f"[WD-EXCLUDE] WD 목록 {len(wd_keys):,}건 로드")

    # ── 원본에서 WD HAWB 전부 제거 ───────────────────────
    df["__HAWB_UPPER__"] = (
        df[col_raw].astype(str).str.strip().str.upper().str.replace("-", "", regex=False)
    )
    before, after = len(df), None
    df = df[~df["__HAWB_UPPER__"].isin(wd_keys)].copy()
    after = len(df)

    print(f"[WD-EXCLUDE] rows {before:,} ▶ {after:,}  (-{before-after:,})")
    return df.drop(columns="__HAWB_UPPER__")



def kpi_calculation(df: pd.DataFrame, cfg: dict):
    # 1) KPI 대상 열 결정
    in_col = cfg.get("in_pbi_column", "In PBI")  # 필요시 "PBI"로 설정해서 넘기면 됨
    if in_col not in df.columns:
        raise KeyError(f"kpi_calculation: '{in_col}' 열을 DataFrame에서 찾을 수 없습니다.")

    # 2) 총 주문 건수 열 (❗ 자동 생성 삭제: 없으면 에러)
    total_col = cfg.get("total_order_col", "Unique Order")
    if total_col not in df.columns:
        raise KeyError(
            f"kpi_calculation: '{total_col}' 열이 없습니다. "
            f"HAWB 중복 처리 단계에서 첫 발생=1, 중복=0으로 '{total_col}'을 생성해 주세요."
        )

    prec = int(cfg.get("precision", 2))

    # 3) PBI 존재 여부(값이 있으면 True)
    s = df[in_col].astype("string")
    has_pbi = s.notna() & s.str.strip().ne("")

    # 첫 발생(=Total Order == 1) 마스크
    first_only = pd.to_numeric(df[total_col], errors="coerce").fillna(0).astype(int).eq(1)

    # ⬇ 변경점: 첫 발생에서만 1 카운트
    df["PBI Count"]    = (has_pbi & first_only).astype(int)

    # 4) 실패 건수·온타임 퍼센트 (행 단위; 분모 0이면 NaN)
    df["Failed Order"] = (first_only.astype(int) - df["PBI Count"]).astype(int)

    # 원래 apply 대신 벡터화 (동일 의미, 더 빠르고 깔끔)
    df["On Time %"] = (df["PBI Count"] * 100).div(
        first_only.astype("Int64").replace({0: pd.NA}).astype("Float64")
    ).round(prec)

    return df


def hawb_validation_debug(df, cfg, hawb_path):
    """
    hawb_validation 결과를 요약 출력해 주는 래퍼
    FG_Match / In PBI 값 분포와 샘플 10행을 콘솔에 찍는다.
    """
    out = hawb_validation(df.copy(), cfg, hawb_path)

    print("\n[hawb_validation_debug] --------")
    print("FG_Match  'Yes'  개수 :", (out["FG_Match"] == "Yes").sum())
    print("In PBI 분포        :", out["In PBI"].value_counts(dropna=False).to_dict())
    print(out[[cfg.get("main_hawb_column", "HAWB / Bill of Lading #"),
            "FG_Match", "In PBI"]].head(10))
    print("---------------------------------\n")
    return out

def file_fill(df: pd.DataFrame, cfg: dict | None = None) -> pd.DataFrame:
    print("file_fill running")
    # 기존 로직을 '동일 의미'로, 컬럼 마스크를 안전하게 계산해 적용
    skip = {"FG", "WD", "PBI", "In PBI", "On Time %", "Remarks"}
    mask = ~df.columns.isin(list(skip))          # True=채워줄 대상 컬럼
    df.loc[:, mask] = df.loc[:, mask].ffill()    # 대상만 ffill
    return df


# Dispatcher
STEP_DISPATCH = {
    "create_columns"  : create_columns,
    "unique_hawb"     : unique_hawb,
    "wd_filtering"    : wd_filtering,
    "remarks_update"  : remarks_update,
    "pbi_multi_merge": pbi_multi_merge,
    "kpi_calculation" : kpi_calculation,
    "sort_hawb": sort_hawb,
    "wd_hawb_exclude": wd_hawb_exclude,
    "remove_by_remarks": remove_by_remarks,
    "file_fill": file_fill,

}

# ──────────────────────────────────────────────
# 2.  JSON 로더 – 국가 폴더만 탐색
# ──────────────────────────────────────────────


def load_step_jsons(country: str, base_dir: str = BASE_DIR):
    """
    ① DATA_DIR(쓰기 가능) → ② TEMPLATE_DIR(읽기 전용) 순으로 검색
    """
    primary   = os.path.join(base_dir, country)
    fallback  = os.path.join(TEMPLATE_DIR, country)

    dir_path  = primary if os.path.isdir(primary) else fallback
    print(" looking for:", dir_path)

    if not os.path.isdir(dir_path):
        raise FileNotFoundError(f"No step folder found for '{country}' -> {dir_path}")

    files = sorted(f for f in os.listdir(dir_path) if f.lower().endswith(".json"))
    return [json.load(open(os.path.join(dir_path, f), encoding="utf-8"))
            for f in files]
            
def _mangle_dupe_cols_like_pandas(cols):
    """
    엑셀 '재읽기'를 거치지 않고 메모리 DF를 바로 받을 때,
    판다스의 mangle_dupe_cols=True 동작(Col, Col.1, Col.2…)을 재현한다.
    """
    seen = {}
    out = []
    for c in map(str, cols):
        if c in seen:
            seen[c] += 1
            out.append(f"{c}.{seen[c]}")  # e.g., "Consignee Name.1"
        else:
            seen[c] = 0
            out.append(c)
    return out

# ──────────────────────────────────────────────
# 3.  통합 실행 함수
# ──────────────────────────────────────────────
def run_validation(
    raw_file_path   : str,
    country         : str,
    hawb_file_path  : str | None = None,
    year            : int | None = None,
    month           : int | None = None,
    base_dir        : str = BASE_DIR,
    df              : pd.DataFrame | None = None
) -> pd.DataFrame:
    """
    전체 파이프라인 실행 후 결과 DataFrame 반환
    - df 가 주어지면 그 DF로 시작(엑셀 재읽기 생략)
    - df 가 None 이면 raw_file_path 를 읽어 시작(예전 동작)
    """
    # ── 0) 입력 소스 결정 ─────────────────────────────────────────
    if df is not None:
        # ★★★ 핵심 방어: 엑셀 재읽기가 해주던 '중복 컬럼 유니크화'를 재현
        df = df.copy()
        df.columns = _mangle_dupe_cols_like_pandas(df.columns)
    else:
        # 예전 방식: 파일에서 읽으면 판다스가 중복 헤더를 자동으로 Col, Col.1…로 보정
        if not raw_file_path:
            raise ValueError("run_validation: df 또는 raw_file_path 중 하나는 반드시 필요합니다.")
        df = pd.read_excel(raw_file_path)

# ── (옵션) 연·월 필터 ─────────────────────────────────────────
    if (year or month):
        # 1) 간단 alias로 컬럼 탐색 (헤더 매핑 없을 때 대비)
        date_aliases = [
            "Declaration Date",            # 기존
            "DECLATION DATE",              # 말레이시아(오타)
            "Custom Form Declaration Date" # 다른 브로커에서 자주 쓰는 표기
        ]
        date_col = next((c for c in date_aliases if c in df.columns), None)

        if date_col:
            col = df[date_col]
            s_str = col.astype("string").str.strip()

            # ── 형식 감지: 우리가 허용한 두 형식('-' 또는 '/')이 다수인지?
            has_dash  = s_str.str.contains("-", na=False)
            has_slash = s_str.str.contains("/", na=False)
            ratio_fmt = float((has_dash | has_slash).mean())  # 구분자 포함 비율

            if ratio_fmt >= 0.6:
                # ── 두 가지 파싱 전략을 시도하고 '유효+그럴듯한' 쪽 선택
                s_default  = pd.to_datetime(col, errors="coerce")
                s_dayfirst = pd.to_datetime(col, errors="coerce", dayfirst=True)

                def plausible_mask(s):
                    y = s.dt.year
                    return (y >= 1990) & (y <= 2100)

                valid_def = s_default.notna() & plausible_mask(s_default)
                valid_df  = s_dayfirst.notna() & plausible_mask(s_dayfirst)

                # 유효치 비율 비교
                ratio_def = float(valid_def.mean())
                ratio_df  = float(valid_df.mean())

                s = s_dayfirst if ratio_df > ratio_def else s_default
                valid = valid_df if ratio_df > ratio_def else valid_def

                # 최소한 하나라도 유효하면 필터 적용, 아니면 안전하게 스킵
                if valid.any():
                    if year:
                        df = df[s.dt.year == year]
                    if month:
                        df = df[s.dt.month == month]
                # else: 모두 비유효 → 형식 불명확으로 판단하고 필터 비적용
            else:
                # 구분자('-' 또는 '/')가 충분히 없으면 (예: '45862') → 필터 비적용
                pass
        # else: 날짜 컬럼을 못 찾으면 기존과 동일하게 필터 건너뜀

    # ── 스텝 JSON 파일 순차 실행 ─────────────────────────────────
    for step_json in load_step_jsons(country, base_dir):
        step_name = step_json["step"]
        conds     = step_json.get("conditions", {})
        func      = STEP_DISPATCH.get(step_name)

        if not func:
            print(f"[WARN] dispatcher 없음: {step_name} -> skip")
            continue

        # --- 스텝별 분기 ---------------------------------------
        if step_name == "pbi_multi_merge":
            if hawb_file_path is None:
                raise ValueError("pbi_multi_merge 단계가 필요하지만 hawb_file_path 가 None")
            df = func(df, conds, pbi_file_path=hawb_file_path)

        elif step_name == "wd_hawb_exclude":              # WD HAWB 완전 제외
            if hawb_file_path is None:
                raise ValueError("wd_hawb_exclude 단계가 필요한데 hawb_file_path 가 None")
            df = func(df, conds, hawb_excel_path=hawb_file_path)

        elif step_name == "hawb_validation":
            if hawb_file_path is None:
                raise ValueError("hawb_validation 단계가 필요한데 hawb_file_path 가 None")
            df = func(df, conds, hawb_excel_path=hawb_file_path)
            hawb_validation_debug(df, conds, hawb_file_path)   # 디버그 출력

        else:  # 일반 스텝
            df = func(df, conds)

        print(f"[DEBUG] after {step_name} -> type(df) = {type(df)}")


    # ── 반드시 DataFrame 반환 ───────────────────────────────────
    return df


# ─────────────────────────────────────────────
# 4.  사용 예 (직접 실행 테스트용)
# ──────────────────────────────────────────────
if __name__ == "__main__":
    validated = run_validation(
        raw_file_path="uploads/Broker_China.xlsx",
        country="China",
        hawb_file_path="uploads/HAWB_Validation_Source.xlsx",
        year=2025, month=6
    )
    validated.to_excel("output/validated_result.xlsx", index=False)
    print("✅ Validation complete -> output/validated_result.xlsx")
