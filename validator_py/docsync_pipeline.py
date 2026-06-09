#!/usr/bin/env python3
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path


def normalize_list_arg(raw: str) -> list[str]:
    return [x.strip() for x in (raw or "").split(",") if x.strip()]


def build_temp_xdrz_from_xdmz(xdmz_files: list[Path]) -> Path:
    temp_dir = Path(tempfile.mkdtemp(prefix="docsync_xdrz_"))
    xdrz_path = temp_dir / "generated_from_xdmz.xdrz"
    with zipfile.ZipFile(xdrz_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for xdmz in xdmz_files:
            zf.write(xdmz, arcname=xdmz.name)
    return xdrz_path


def collect_xdmz_files(xdmz_dir: Path, xdmz_filter: list[str]) -> list[Path]:
    all_files = sorted(xdmz_dir.glob("*.xdmz"), key=lambda p: p.name.upper())
    if not xdmz_filter:
        return all_files
    wanted = {x.upper() for x in xdmz_filter}
    out = [p for p in all_files if p.name.upper() in wanted or p.stem.upper() in wanted]
    return out


def run_command(cmd: list[str]) -> None:
    proc = subprocess.run(cmd, check=False)
    if proc.returncode != 0:
        raise SystemExit(proc.returncode)


def main() -> None:
    ap = argparse.ArgumentParser(description="Pipeline unificado Reference Equity DocSync")
    ap.add_argument("--doc", required=True, help="Caminho para o DOCX publicado.")
    ap.add_argument("--out", required=True, help="Diretório de saída.")
    ap.add_argument("--xdrz", default="", help="Entrada XDRZ opcional.")
    ap.add_argument("--xdmz-dir", default="", help="Pasta opcional com arquivos .xdmz.")
    ap.add_argument("--xdmz", default="", help="Lista opcional, separada por vírgula, de nomes ou stems de arquivos .xdmz.")
    ap.add_argument("--enrich-web", action="store_true")
    ap.add_argument("--enrich-cache", default="")
    ap.add_argument("--enrich-timeout-sec", type=int, default=12)
    ap.add_argument("--scope-mode", choices=["anchor-window", "name-match"], default="anchor-window")
    ap.add_argument("--write-scope-debug-json", action="store_true")
    ap.add_argument("--extractor", default="")
    ap.add_argument("--include-pass-in-dashboard", action="store_true")
    ap.add_argument("--mapping-file", default="")
    ap.add_argument("--flex-setup-file", default="")
    ap.add_argument("--skip-web-post", action="store_true", help="Ignora consulta web na etapa de pós-processamento.")
    args = ap.parse_args()

    root = Path(__file__).resolve().parent.parent
    compare_script = root / "validator_py" / "docsync_compare.py"
    fill_script = root / "validator_py" / "fill_dashboard_descriptions.py"

    doc_path = Path(args.doc).resolve()
    out_dir = Path(args.out).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    if doc_path.suffix.lower() != ".docx":
        ap.error(f"Formato de documentação inválido: {doc_path}. Use .docx.")
    if not doc_path.exists():
        ap.error(f"DOCX não encontrado: {doc_path}")

    xdrz_input = Path(args.xdrz).resolve() if args.xdrz.strip() else None
    xdmz_dir = Path(args.xdmz_dir).resolve() if args.xdmz_dir.strip() else None
    xdmz_filter = normalize_list_arg(args.xdmz)

    temp_xdrz: Path | None = None
    temp_dir_to_cleanup: Path | None = None

    if xdrz_input:
        if not xdrz_input.exists():
            ap.error(f"XDRZ não encontrado: {xdrz_input}")
        effective_xdrz = xdrz_input
    else:
        if not xdmz_dir:
            ap.error("Informe --xdrz ou --xdmz-dir.")
        if not xdmz_dir.exists():
            ap.error(f"Diretório XDMZ não encontrado: {xdmz_dir}")
        xdmz_files = collect_xdmz_files(xdmz_dir, xdmz_filter)
        if not xdmz_files:
            ap.error(f"Nenhum arquivo .xdmz encontrado em: {xdmz_dir}")
        temp_xdrz = build_temp_xdrz_from_xdmz(xdmz_files)
        temp_dir_to_cleanup = temp_xdrz.parent
        effective_xdrz = temp_xdrz

    try:
        compare_cmd = [
            sys.executable,
            str(compare_script),
            "--xdrz",
            str(effective_xdrz),
            "--doc",
            str(doc_path),
            "--out",
            str(out_dir),
            "--enrich-timeout-sec",
            str(max(args.enrich_timeout_sec, 3)),
            "--scope-mode",
            args.scope_mode,
        ]
        if args.enrich_web:
            compare_cmd.append("--enrich-web")
        if args.enrich_cache.strip():
            compare_cmd.extend(["--enrich-cache", str(Path(args.enrich_cache).resolve())])
        if args.write_scope_debug_json:
            compare_cmd.append("--write-scope-debug-json")
        if args.extractor.strip():
            compare_cmd.extend(["--extractor", args.extractor.strip()])
        if args.include_pass_in_dashboard:
            compare_cmd.append("--include-pass-in-dashboard")

        print("[PIPELINE] Etapa 1/2 - comparação")
        run_command(compare_cmd)

        dashboard_path = out_dir / "comparison_dashboard.html"
        mapping_file = Path(args.mapping_file).resolve() if args.mapping_file.strip() else (root / "assets" / "table_doc_mapping.txt")
        flex_setup_file = Path(args.flex_setup_file).resolve() if args.flex_setup_file.strip() else (root / "assets" / "G_FLEX_SETUP.xlsx")

        fill_cmd = [
            sys.executable,
            str(fill_script),
            "--input",
            str(dashboard_path),
            "--output",
            str(dashboard_path),
            "--mapping-file",
            str(mapping_file),
            "--flex-setup-file",
            str(flex_setup_file),
        ]
        if args.skip_web_post:
            fill_cmd.append("--skip-web")

        print("[PIPELINE] Etapa 2/2 - pós-processamento do dashboard")
        run_command(fill_cmd)

        print("[PIPELINE] Concluído")
        print(f"[PIPELINE] Dashboard: {dashboard_path}")
        print(f"[PIPELINE] Relatório JSON: {out_dir / 'comparison_report.json'}")
        print(f"[PIPELINE] Relatório MD: {out_dir / 'comparison_report.md'}")
    finally:
        if temp_dir_to_cleanup and temp_dir_to_cleanup.exists():
            shutil.rmtree(temp_dir_to_cleanup, ignore_errors=True)


if __name__ == "__main__":
    main()

