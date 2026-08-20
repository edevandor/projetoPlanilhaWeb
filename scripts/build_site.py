#!/usr/bin/env python3
"""Gera o site estático da aba MOSTRUARIO da planilha de preços."""

from __future__ import annotations

import argparse
import html
import os
import shutil
import subprocess
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

SHEET_NAME = "MOSTRUARIO"
MAX_COLUMNS = 12
MAX_SCAN_ROWS = 10_000
EMPTY_ROWS_TO_STOP = 25
PRICE_COLUMNS = {8, 9, 10, 11}  # I:L, índices 1-based


def recalculate_workbook(source: Path, workdir: Path) -> Path:
    """Recalcula o XLSX com LibreOffice e retorna o arquivo recalculado."""
    if shutil.which("soffice") is None:
        raise RuntimeError(
            "LibreOffice (soffice) é obrigatório para resolver as fórmulas da planilha."
        )

    recalculated_dir = workdir / "recalculated"
    recalculated_dir.mkdir()
    copied_source = workdir / source.name
    shutil.copy2(source, copied_source)

    env = os.environ.copy()
    env["HOME"] = str(workdir / "home")
    Path(env["HOME"]).mkdir()

    command = [
        "soffice",
        "--headless",
        "--convert-to",
        "xlsx",
        "--outdir",
        str(recalculated_dir),
        str(copied_source),
    ]
    result = subprocess.run(command, capture_output=True, text=True, env=env)
    recalculated = recalculated_dir / source.name
    if result.returncode != 0 or not recalculated.exists():
        raise RuntimeError(
            "Falha ao recalcular a planilha com LibreOffice.\n"
            f"stdout: {result.stdout}\n"
            f"stderr: {result.stderr}"
        )
    return recalculated


def read_mostruario(workbook_path: Path) -> tuple[list[str], list[list[Any]]]:
    workbook = load_workbook(workbook_path, data_only=True)
    if SHEET_NAME not in workbook.sheetnames:
        raise ValueError(f"A aba obrigatória {SHEET_NAME!r} não existe na planilha.")

    worksheet = workbook[SHEET_NAME]
    headers = [worksheet.cell(1, column).value for column in range(1, MAX_COLUMNS + 1)]
    headers = [str(value or "") for value in headers]

    rows: list[list[Any]] = []
    empty_streak = 0
    scan_limit = min(worksheet.max_row, MAX_SCAN_ROWS)
    for row_number in range(2, scan_limit + 1):
        values = [worksheet.cell(row_number, column).value for column in range(1, MAX_COLUMNS + 1)]
        if not any(value not in (None, "") for value in values):
            empty_streak += 1
            if rows and empty_streak >= EMPTY_ROWS_TO_STOP:
                break
            continue
        empty_streak = 0
        rows.append(values)

    if not rows:
        raise ValueError(f"A aba {SHEET_NAME!r} não contém registros de produtos.")
    if len(rows) >= scan_limit and scan_limit == MAX_SCAN_ROWS:
        raise ValueError(
            f"A leitura atingiu o limite de {MAX_SCAN_ROWS} linhas; revise o contrato da aba."
        )
    return headers, rows


def format_value(value: Any, column_number: int) -> str:
    if value is None:
        return ""
    if column_number in PRICE_COLUMNS and isinstance(value, (int, float)):
        return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return str(value)


def render_html(headers: list[str], rows: list[list[Any]], source_name: str) -> str:
    generated_at = datetime.now(timezone.utc).astimezone().strftime("%d/%m/%Y %H:%M:%S %Z")
    body_rows = []
    for row in rows:
        search_text = " ".join(str(value or "") for value in row).lower()
        cells = []
        for position, value in enumerate(row, start=1):
            css_class = " class=\"price\"" if position in PRICE_COLUMNS else ""
            cells.append(f"<td{css_class}>{html.escape(format_value(value, position))}</td>")
        rendered_cells = "".join(cells)
        body_rows.append(
            f'<tr data-search="{html.escape(search_text, quote=True)}">{rendered_cells}</tr>'
        )

    table_header = "".join(f"<th>{html.escape(header)}</th>" for header in headers)
    table_body = "\n".join(body_rows)
    escaped_source = html.escape(source_name)

    return f"""<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Mostruário de preços</title>
  <style>
    :root {{ color-scheme: light; font-family: Inter, system-ui, -apple-system, sans-serif; }}
    body {{ margin: 0; background: #f4f6f8; color: #18212b; }}
    header {{ padding: 28px clamp(16px, 4vw, 56px) 20px; background: #102a43; color: white; }}
    h1 {{ margin: 0 0 8px; font-size: clamp(1.5rem, 3vw, 2.2rem); }}
    .meta {{ margin: 0; opacity: .8; font-size: .9rem; }}
    main {{ padding: 24px clamp(16px, 4vw, 56px) 48px; }}
    .toolbar {{ display: flex; gap: 12px; align-items: center; margin-bottom: 16px; flex-wrap: wrap; }}
    input {{ width: min(560px, 100%); padding: 12px 14px; border: 1px solid #bcccdc; border-radius: 8px; font-size: 1rem; }}
    #count {{ color: #52606d; font-size: .9rem; }}
    .table-wrap {{ overflow: auto; background: white; border: 1px solid #d9e2ec; border-radius: 10px; box-shadow: 0 2px 8px #102a4312; }}
    table {{ border-collapse: collapse; width: 100%; min-width: 980px; }}
    th, td {{ padding: 11px 12px; border-bottom: 1px solid #e6edf3; text-align: left; white-space: nowrap; }}
    th {{ position: sticky; top: 0; background: #eaf2f8; color: #243b53; font-size: .82rem; text-transform: uppercase; letter-spacing: .03em; }}
    td.price {{ text-align: right; font-variant-numeric: tabular-nums; }}
    tbody tr:hover {{ background: #f0f7ff; }}
    tbody tr:last-child td {{ border-bottom: 0; }}
    .empty {{ padding: 28px; color: #52606d; }}
  </style>
</head>
<body>
  <header>
    <h1>Mostruário de preços</h1>
    <p class="meta">Fonte: {escaped_source} · Gerado em {html.escape(generated_at)}</p>
  </header>
  <main>
    <div class="toolbar">
      <input id="search" type="search" placeholder="Buscar por código, SKU, modelo, marca..." autocomplete="off">
      <span id="count"></span>
    </div>
    <div class="table-wrap">
      <table>
        <thead><tr>{table_header}</tr></thead>
        <tbody id="products">{table_body}</tbody>
      </table>
      <div id="empty" class="empty" hidden>Nenhum produto encontrado.</div>
    </div>
  </main>
  <script>
    const search = document.querySelector('#search');
    const rows = [...document.querySelectorAll('#products tr')];
    const count = document.querySelector('#count');
    const empty = document.querySelector('#empty');
    function update() {{
      const query = search.value.trim().toLowerCase();
      let visible = 0;
      for (const row of rows) {{
        const show = !query || row.dataset.search.includes(query);
        row.hidden = !show;
        if (show) visible++;
      }}
      count.textContent = `${{visible}} de ${{rows.length}} produtos`;
      empty.hidden = visible !== 0;
    }}
    search.addEventListener('input', update);
    update();
  </script>
</body>
</html>
"""


def build(input_path: Path, output_dir: Path) -> int:
    input_path = input_path.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="planilia-web-") as temporary:
        recalculated = recalculate_workbook(input_path, Path(temporary))
        headers, rows = read_mostruario(recalculated)

    output_path = output_dir / "index.html"
    output_path.write_text(render_html(headers, rows, input_path.name), encoding="utf-8")
    print(f"Gerado: {output_path}")
    print(f"Aba: {SHEET_NAME}")
    print(f"Produtos: {len(rows)}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, required=True, help="Caminho do XLSX fonte")
    parser.add_argument("--output", type=Path, required=True, help="Diretório do site gerado")
    args = parser.parse_args()
    return build(args.input, args.output)


if __name__ == "__main__":
    raise SystemExit(main())
