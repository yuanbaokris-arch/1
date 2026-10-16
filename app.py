"""公交营运预算试算：原始工作簿、真实重算和不可覆盖的版本。"""
from __future__ import annotations

from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
import hashlib
import json
import math
import os
import shutil
import threading
import uuid
from zipfile import BadZipFile, ZipFile

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from starlette.middleware.gzip import GZipMiddleware
from pydantic import BaseModel, Field

from budget_engine import compare_formula_caches, parse_workbook, recalculate_preserving_formulas
from budget_export import add_monthly_report


ROOT = Path(__file__).resolve().parent
DATA_ROOT = Path(os.environ.get("BUS_BUDGET_DATA_DIR", str(ROOT / "data"))).resolve()
DEMO_PATH = Path(os.environ.get(
    "BUS_BUDGET_DEMO_PATH",
    str(ROOT / "example.xlsx"),
))
MAX_UPLOAD = 24 * 1024 * 1024
app = FastAPI(title="公交营运预算调整工具", docs_url="/api/docs", redoc_url=None)
app.add_middleware(GZipMiddleware, minimum_size=1000)
_mutation_lock = threading.Lock()


def now():
    return datetime.now(timezone.utc).isoformat()


def write_json(path: Path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(path.name + "." + uuid.uuid4().hex + ".tmp")
    temp.write_text(json.dumps(value, ensure_ascii=False, allow_nan=False), encoding="utf-8")
    os.replace(temp, path)


def read_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


@contextmanager
def changing():
    if not _mutation_lock.acquire(blocking=False):
        raise HTTPException(409, "正在导入或重算，请等待当前操作完成。")
    try:
        yield
    finally:
        _mutation_lock.release()


def dataset_dir():
    active = DATA_ROOT / "active.json"
    if not active.exists():
        raise HTTPException(404, "请先导入预算基础表。")
    dataset_id = read_json(active)["id"]
    return DATA_ROOT / "datasets" / dataset_id


def metadata():
    directory = dataset_dir()
    return directory, read_json(directory / "dataset.json")


def version_data(directory, meta, version_id=None):
    versions = meta["versions"]
    selected = next((v for v in versions if v["id"] == version_id), None) if version_id else versions[-1]
    if selected is None:
        raise HTTPException(404, "找不到这个已保存版本。")
    return read_json(directory / "versions" / (selected["id"] + ".json"))


def draft_data(directory):
    path = directory / "draft.json"
    return read_json(path) if path.exists() else None


def state(version_id=None):
    capabilities = {"demo_available": DEMO_PATH.is_file()}
    if not (DATA_ROOT / "active.json").exists():
        return {"dataset": None, "version": None, "draft": None, "capabilities": capabilities}
    directory, meta = metadata()
    return {"dataset": meta, "version": version_data(directory, meta, version_id), "draft": draft_data(directory), "capabilities": capabilities}


def numeric(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def normalize_warnings(items, original=None):
    def fingerprint(x):
        return tuple(str(x.get(k, "")) for k in ("code", "sheet", "cell", "message"))
    originals = {fingerprint(x) for x in original} if original is not None else None
    result = []
    for item in items:
        w = dict(item) if isinstance(item, dict) else {"message": str(item)}
        w.setdefault("code", "WORKBOOK_NOTICE")
        w.setdefault("severity", "warning")
        w.setdefault("sheet", "")
        w.setdefault("cell", "")
        w["origin"] = "existing" if originals is None or fingerprint(w) in originals else "new"
        result.append(w)
    return result


def to_view(parsed, originals=None, original_warnings=None):
    original_rows = {r["key"]: r for r in originals or []}
    rows = []
    warnings = list(parsed.get("warnings", []))
    for month in parsed["months"]:
        warnings.extend(month.get("warnings", []))
        for route in month["routes"]:
            warnings.extend(route.get("warnings", []))
            row = {
                "key": route["id"], "month": month["name"],
                "route": route["line"], "department": route["department"],
                "fleet": route["team"], "source_row": route["row"],
                "summary_row": route["summary_row"],
                "vehicles": route["vehicles_a"], "day_trips": route["plan_trips"]["A"],
                "a_days": route["a_days"], "monthly_trips": route["monthly_trips"],
                "max_vehicles": route["max_regular"], "reserve": route["reserve_a"],
                "trips_per_vehicle": route["trips_per_vehicle_a"],
                "daily": route.get("daily", []), "calendar_days": month["days"],
            }
            original = original_rows.get(row["key"], row)
            for k in ("vehicles", "day_trips", "monthly_trips", "max_vehicles", "reserve"):
                row["original_" + k] = original[k]
            row["trip_delta"] = row["monthly_trips"] - row["original_monthly_trips"]
            rows.append(row)
    total = sum(r["monthly_trips"] for r in rows)
    original_total = sum(r["original_monthly_trips"] for r in rows)
    return {
        "rows": rows,
        "error_cells": parsed.get("error_cells", []),
        "totals": {"monthly_trips": total, "original_monthly_trips": original_total,
                   "trip_delta": total - original_total,
                   "affected_rows": sum(r["vehicles"] != r["original_vehicles"] for r in rows)},
        "warnings": normalize_warnings(warnings, original_warnings),
    }


def validate_archive(path):
    try:
        with ZipFile(path) as z:
            infos = z.infolist()
            if len(infos) > 4000 or sum(i.file_size for i in infos) > 240 * 1024 * 1024:
                raise ValueError("工作簿内容过大，请使用原始预算基础表模板。")
            if any(i.flag_bits & 1 for i in infos):
                raise ValueError("暂不支持加密工作簿，请上传未加密的副本。")
            if "xl/workbook.xml" not in z.namelist():
                raise ValueError("这不是有效的 .xlsx 工作簿。")
            if any(n.startswith("xl/externalLinks/") or n.endswith("vbaProject.bin") for n in z.namelist()):
                raise ValueError("此版本不接受含外部链接或宏的工作簿，以免试算依赖未提供的数据。")
            if z.testzip() is not None:
                raise ValueError("工作簿压缩内容损坏。")
    except BadZipFile as e:
        raise ValueError("无法读取文件，请上传有效的 .xlsx 文件。") from e


def import_file(content, filename):
    if not filename.lower().endswith(".xlsx"):
        raise HTTPException(422, "请上传 .xlsx 格式的预算基础表。")
    if not content or len(content) > MAX_UPLOAD:
        raise HTTPException(422, "文件为空或超过24MB。")
    directory = DATA_ROOT / "datasets" / uuid.uuid4().hex
    directory.mkdir(parents=True)
    original = directory / "original.xlsx"
    original.write_bytes(content)
    try:
        validate_archive(original)
        parsed = parse_workbook(original)
        view = to_view(parsed)
        # 用真实引擎空操作重算，检验缓存而不覆盖 V0。
        baseline = directory / "baseline-calculated.xlsx"
        calculation = recalculate_preserving_formulas(original, [], baseline)
        formula_check = compare_formula_caches(original, baseline)
        baseline_view = to_view(parse_workbook(baseline))
        differences = []
        baseline_rows = {r["key"]: r for r in baseline_view["rows"]}
        for row in view["rows"]:
            other = baseline_rows[row["key"]]
            for k in ("vehicles", "day_trips", "a_days", "monthly_trips", "max_vehicles", "reserve"):
                a, b = row[k], other[k]
                if a != b and not (numeric(a) and numeric(b) and math.isclose(a, b, abs_tol=1e-7, rel_tol=1e-12)):
                    differences.append({"key": row["key"], "field": k, "original": a, "recalculated": b})
            if row["daily"] != other["daily"]:
                differences.append({"key": row["key"], "field": "daily"})
        verified = not differences and formula_check["identical"]
        if not verified:
            view["warnings"].append({"code": "BASELINE_MISMATCH", "severity": "error", "origin": "existing",
                "sheet": "", "cell": "", "message": "空操作重算与原文件结果存在差异，已保留V0；暂不能调整，请先核对差异。"})
        version_id = uuid.uuid4().hex
        (directory / "versions").mkdir()
        dest = directory / "versions" / (version_id + ".xlsx")
        shutil.copyfile(original, dest)
        os.chmod(original, 0o444)
        os.chmod(dest, 0o444)
        version = {"id": version_id, "label": "V0", "parent_id": None, "created_at": now(),
                   "reason": "导入原始预算基础表，不可覆盖", "operations": [], **view}
        write_json(directory / "versions" / (version_id + ".json"), version)
        meta = {"id": directory.name, "name": Path(filename).name, "sha256": hashlib.sha256(content).hexdigest(),
                "created_at": version["created_at"], "route_count": len({r["route"] for r in view["rows"]}),
                "months": [m["name"] for m in parsed["months"]], "warnings": view["warnings"],
                "baseline_verified": verified, "baseline_differences": differences[:100],
                "baseline_formula_check": {**formula_check, "differences": formula_check["differences"][:100]},
                "baseline_engine": calculation.get("engine", "LibreOffice") if isinstance(calculation, dict) else "LibreOffice",
                "versions": [{k: version[k] for k in ("id", "label", "parent_id", "created_at", "reason")}]}
        write_json(directory / "dataset.json", meta)
        write_json(DATA_ROOT / "active.json", {"id": directory.name})
        return state()
    except Exception:
        shutil.rmtree(directory, ignore_errors=True)
        raise


def operation_failure(error):
    if isinstance(error, HTTPException):
        raise error
    if isinstance(error, (ValueError, KeyError, BadZipFile)):
        raise HTTPException(422, "无法完成操作：" + str(error)) from error
    if isinstance(error, (RuntimeError, TimeoutError)):
        raise HTTPException(503, "公式重算未完成，当前版本未改变。" + str(error)) from error
    raise error


class DraftRequest(BaseModel):
    version_id: str
    discard: bool = False


class AdjustRequest(BaseModel):
    keys: list[str] = Field(min_length=1, max_length=5000)
    delta: int = -1
    reason: str = Field(default="", max_length=1000)
    locked_routes: list[str] = Field(default_factory=list, max_length=1000)


class SaveRequest(BaseModel):
    reason: str = Field(min_length=1, max_length=1000)


@app.get("/api/health")
def health():
    return {"status": "ok", "app": "公交营运预算调整", "calculation": "保留原公式，使用LibreOffice重算",
            "instance_id": hashlib.sha256(str(DATA_ROOT).encode()).hexdigest()[:16]}


@app.get("/api/state")
def get_state(version_id: str | None = None):
    return state(version_id)


@app.post("/api/import")
def upload(file: UploadFile = File(...)):
    with changing():
        try:
            return import_file(file.file.read(MAX_UPLOAD + 1), file.filename or "")
        except Exception as e:
            operation_failure(e)


@app.post("/api/demo")
def demo():
    with changing():
        if not DEMO_PATH.is_file():
            raise HTTPException(404, "当前环境中没有示例基础表，请使用导入按钮上传文件。")
        try:
            return import_file(DEMO_PATH.read_bytes(), DEMO_PATH.name)
        except Exception as e:
            operation_failure(e)


@app.post("/api/draft")
def new_draft(request: DraftRequest):
    with changing():
        directory, meta = metadata()
        if not meta["baseline_verified"]:
            raise HTTPException(409, "空操作重算对照尚未通过，请先核对原表结果差异。")
        parent = version_data(directory, meta, request.version_id)
        old = draft_data(directory)
        if old and old["parent_id"] == parent["id"]:
            return state(parent["id"])
        if old and old["operations"] and not request.discard:
            raise HTTPException(409, "还有未保存的草稿。请先保存，或确认放弃草稿后切换。")
        if old:
            write_json(directory / "discarded-drafts" / (old["id"] + ".json"), {**old, "discarded_at": now()})
        draft = {**parent, "id": uuid.uuid4().hex, "label": "未保存草稿", "parent_id": parent["id"],
                 "created_at": now(), "reason": "", "operations": [], "changes": [], "undo_stack": [],
                 "audit_history": [],
                 "workbook": str(directory / "versions" / (parent["id"] + ".xlsx"))}
        write_json(directory / "draft.json", draft)
        return state(parent["id"])


def apply_changes(directory, meta, draft, changes):
    parent = version_data(directory, meta, draft["parent_id"])
    baseline = version_data(directory, meta, meta["versions"][0]["id"])
    if not changes:
        draft.update({k: parent[k] for k in ("rows", "totals", "warnings", "error_cells")})
        draft["workbook"] = str(directory / "versions" / (parent["id"] + ".xlsx"))
    else:
        temp_dir = directory / "draft-files"
        temp_dir.mkdir(exist_ok=True)
        dest = temp_dir / (uuid.uuid4().hex + ".xlsx")
        try:
            recalculate_preserving_formulas(directory / "versions" / (parent["id"] + ".xlsx"), changes, dest)
            view = to_view(parse_workbook(dest), baseline["rows"], baseline["warnings"])
            original_errors = {(e["sheet"], e["cell"], str(e["value"])) for e in baseline.get("error_cells", [])}
            new_errors = [e for e in view["error_cells"] if (e["sheet"], e["cell"], str(e["value"])) not in original_errors]
            if new_errors:
                error = new_errors[0]
                raise ValueError(f'重算产生新的公式错误 {error["sheet"]}!{error["cell"]}={error["value"]}，本次操作未应用。')
            formula_errors = [w for w in view["warnings"] if w["origin"] == "new" and w["code"] in ("FORMULA_ERROR", "CELL_ERROR", "formula_errors")]
            if formula_errors:
                raise ValueError("重算产生新的公式错误，未应用本次操作：" + formula_errors[0]["message"])
            draft.update(view)
            draft["workbook"] = str(dest)
        except Exception:
            dest.unlink(missing_ok=True)
            raise
    draft["changes"] = changes
    return draft


def adjust(request, restoring=False):
    directory, meta = metadata()
    draft = draft_data(directory)
    if draft is None:
        raise HTTPException(409, "请先从一个已保存版本建立草稿。")
    if request.delta != -1:
        raise HTTPException(422, "第一版仅支持A计划营运1投放减1；局部恢复请使用恢复操作。")
    parent = version_data(directory, meta, draft["parent_id"])
    parent_rows = {r["key"]: r for r in parent["rows"]}
    current_rows = {r["key"]: r for r in draft["rows"]}
    changes = {(c["sheet"], c["cell"]): c for c in draft["changes"]}
    scope = []
    for key in dict.fromkeys(request.keys):
        row = current_rows.get(key)
        if row is None:
            raise HTTPException(422, "所选记录不存在，请刷新页面后重试。")
        if row["route"] in request.locked_routes:
            raise HTTPException(422, row["route"] + "已锁定，本次操作未应用。")
        before = row["vehicles"]
        if not numeric(before) or int(before) != before or before < 0:
            raise HTTPException(422, f'{row["month"]} {row["route"]}当前投放数不是非负整数，请先核对原始数据。本批次未作修改。')
        after = parent_rows[key]["vehicles"] if restoring else before - 1
        if not numeric(after) or after < 0 or int(after) != after:
            raise HTTPException(422, f'{row["month"]} {row["route"]}不能减车：投放数必须为非负整数。本批次未作修改。')
        if after == before:
            continue
        cell = "J" + str(row["source_row"])
        change_key = (row["month"], cell)
        if after == parent_rows[key]["vehicles"]:
            changes.pop(change_key, None)
        else:
            changes[change_key] = {"sheet": row["month"], "cell": cell, "value": int(after)}
        scope.append({"key": key, "route": row["route"], "month": row["month"], "cell": cell,
                      "before": before, "after": after})
    if not scope:
        raise HTTPException(422, "所选记录已与父版本相同，无需恢复。")
    old = {"changes": draft["changes"], "operation_count": len(draft["operations"])}
    before_total = draft["totals"]["monthly_trips"]
    updated = apply_changes(directory, meta, draft, list(changes.values()))
    updated["undo_stack"].append(old)
    updated["operations"].append({"id": uuid.uuid4().hex, "type": "restore" if restoring else "reduce",
                                  "created_at": now(), "actor": "当前使用者", "reason": request.reason.strip(),
                                  "scope": scope, "count": len(scope), "plan": "A", "field": "营运1投放车辆数",
                                  "trip_delta": updated["totals"]["monthly_trips"] - before_total})
    write_json(directory / "draft.json", updated)
    return state(updated["parent_id"])


@app.post("/api/draft/adjust")
def reduce_vehicles(request: AdjustRequest):
    with changing():
        try:
            return adjust(request)
        except Exception as e:
            operation_failure(e)


@app.post("/api/draft/restore")
def restore_vehicles(request: AdjustRequest):
    with changing():
        try:
            return adjust(request, restoring=True)
        except Exception as e:
            operation_failure(e)


@app.post("/api/draft/undo")
def undo():
    with changing():
        directory, meta = metadata()
        draft = draft_data(directory)
        if not draft or not draft["undo_stack"]:
            raise HTTPException(422, "没有可以撤销的草稿操作。")
        checkpoint = draft["undo_stack"][-1]
        try:
            updated = apply_changes(directory, meta, draft, checkpoint["changes"])
        except Exception as e:
            operation_failure(e)
        updated["undo_stack"].pop()
        # 保留被撤销操作的记录，使版本保存后仍能追溯全过程。
        removed = updated["operations"][checkpoint["operation_count"]:]
        updated.setdefault("audit_history", []).extend({**op, "undone_at": now()} for op in removed)
        updated["operations"] = updated["operations"][:checkpoint["operation_count"]]
        write_json(directory / "draft.json", updated)
        return state(updated["parent_id"])


@app.post("/api/draft/discard")
def discard():
    with changing():
        directory, _ = metadata()
        draft = draft_data(directory)
        if draft:
            archive = directory / "discarded-drafts" / (draft["id"] + ".json")
            write_json(archive, {**draft, "discarded_at": now()})
            (directory / "draft.json").unlink()
        return state()


@app.post("/api/draft/save")
def save(request: SaveRequest):
    reason = request.reason.strip()
    if not reason:
        raise HTTPException(422, "请填写版本调整原因。")
    with changing():
        directory, meta = metadata()
        draft = draft_data(directory)
        if not draft or not draft["operations"]:
            raise HTTPException(422, "草稿还没有操作记录，无需另存版本。")
        version_id = uuid.uuid4().hex
        version = {k: v for k, v in draft.items() if k not in ("undo_stack", "workbook")}
        version.update({"id": version_id, "label": "V" + str(len(meta["versions"])), "reason": reason, "created_at": now()})
        workbook = directory / "versions" / (version_id + ".xlsx")
        shutil.copyfile(draft["workbook"], workbook)
        os.chmod(workbook, 0o444)
        write_json(directory / "versions" / (version_id + ".json"), version)
        meta["versions"].append({k: version[k] for k in ("id", "label", "parent_id", "created_at", "reason")})
        write_json(directory / "dataset.json", meta)
        (directory / "draft.json").unlink()
        return state(version_id)


@app.get("/api/export")
def export(version_id: str, kind: str = "monthly"):
    if kind not in ("monthly", "workbook"):
        raise HTTPException(422, "请选择月趟次表或完整工作簿。")
    directory, meta = metadata()
    version = version_data(directory, meta, version_id)
    source = directory / "versions" / (version_id + ".xlsx")
    name = f'2027年营运预算_{version["label"]}'
    if kind == "workbook":
        return FileResponse(source, filename=name + "_原公式工作簿.xlsx",
                            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    # 导出只读取不可覆盖的版本；临时产物不改变当前草稿。
    destination = directory / "exports" / (version_id + "-monthly.xlsx")
    destination.parent.mkdir(exist_ok=True)
    if not destination.exists():
        temp = destination.with_name(uuid.uuid4().hex + ".xlsx")
        try:
            add_monthly_report(source, temp, version["label"], version_id, version["reason"], meta["sha256"])
            os.replace(temp, destination)
        except Exception as e:
            temp.unlink(missing_ok=True)
            operation_failure(e)
    return FileResponse(destination, filename=name + "_月趟次试算.xlsx",
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/", include_in_schema=False)
def index():
    return FileResponse(ROOT / "static" / "index.html")


app.mount("/static", StaticFiles(directory=ROOT / "static"), name="assets")
