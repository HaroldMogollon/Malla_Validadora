from __future__ import annotations

import io
import re
import unicodedata
from datetime import date, datetime
from difflib import SequenceMatcher

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Malla de validación", page_icon="✅", layout="wide")


def norm(value: object) -> str:
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    return re.sub(r"[^A-Z0-9]+", " ", text.upper()).strip()


def read_sheets(upload) -> dict[str, pd.DataFrame]:
    upload.seek(0)
    return pd.read_excel(upload, sheet_name=None, header=None, dtype=object, engine="openpyxl")


def detect_header_row(raw: pd.DataFrame) -> int:
    for i, row in raw.iterrows():
        nonempty = [str(x).strip() for x in row.tolist() if pd.notna(x) and str(x).strip()]
        if len(nonempty) >= 2:
            return int(i)
    return 0


def parse_rules(raw: pd.DataFrame) -> tuple[list[dict], int]:
    header_row = 0
    for i, row in raw.iterrows():
        vals = [norm(x) for x in row.tolist() if pd.notna(x)]
        if any("TIPO CAMPO" in x for x in vals) and any("OBLIGATORIO" in x for x in vals):
            header_row = int(i)
            break
        if any("COLUMNA" in x or "CAMPO" in x for x in vals) and any("DESCRIPCION" in x or "REGLA" in x for x in vals):
            header_row = int(i)
            break
    header_values = [norm(x) for x in raw.iloc[header_row].tolist()]
    structured = any("TIPO CAMPO" in x for x in header_values) and any("OBLIGATORIO" in x for x in header_values)
    if structured:
        def find_col(*needles: str):
            return next((i for i, val in enumerate(header_values) if any(needle in val for needle in needles)), None)
        field_i = find_col("CAMPO DE REGLA", "COLUMNA", "CAMPO")
        type_i = find_col("TIPO CAMPO", "TIPO")
        required_i = find_col("OBLIGATORIO")
        na_i = find_col("NO APLICA")
        accepted_i = find_col("VALORES ACEPTADOS", "VALORES PERMITIDOS")
        parsed = []
        for _, row in raw.iloc[header_row + 1 :].iterrows():
            field = row.iloc[field_i] if field_i is not None else None
            if pd.isna(field) or not str(field).strip():
                continue
            kind = str(row.iloc[type_i]).strip() if type_i is not None and pd.notna(row.iloc[type_i]) else "Sin validación de tipo"
            required_val = row.iloc[required_i] if required_i is not None else False
            na_val = row.iloc[na_i] if na_i is not None else None
            accepted_val = row.iloc[accepted_i] if accepted_i is not None else None
            description = "" if pd.isna(accepted_val) else str(accepted_val).strip()
            parsed.append({
                "field": str(field).strip(), "description": description,
                "declared_type": kind, "required_default": bool(required_val) and norm(required_val) not in {"FALSE", "0", "NO"},
                "no_aplica_default": norm(na_val) == "X",
                "allowed_default": "\n".join(x.strip() for x in description.splitlines() if x.strip() and not norm(x).startswith("NO DEBE SER NULO")),
                "max_date": extract_max_date(description), "structured": True,
            })
        return parsed, header_row
    first = next((i for i, x in enumerate(raw.iloc[header_row].tolist()) if "COLUMNA" in norm(x) or "CAMPO" in norm(x)), 0)
    second = next((i for i, x in enumerate(raw.iloc[header_row].tolist()) if ("DESCRIPCION" in norm(x) or "REGLA" in norm(x)) and i != first), min(first + 1, raw.shape[1] - 1))
    result: list[dict] = []
    active = None
    for _, row in raw.iloc[header_row + 1 :].iterrows():
        field = row.iloc[first] if first < len(row) else None
        desc = row.iloc[second] if second < len(row) else None
        field_text = "" if pd.isna(field) else str(field).strip()
        desc_text = "" if pd.isna(desc) else str(desc).strip()
        if field_text:
            # Some rule books place the description on the next row; skip only
            # obvious section/sub-label rows that do not represent data fields.
            section = norm(field_text)
            if not desc_text and (section in {"ADICION", "DEL CONTRATO", "VALOR"} or section.startswith("DESCRIPCION")):
                active = None
                continue
            active = {"field": field_text, "description": desc_text, "structured": False}
            result.append(active)
        elif active and desc_text:
            active["description"] += "\n" + desc_text
    return result, header_row


def extract_max_date(description: str):
    match = re.search(r"(?:MAYOR\s+A|POSTERIOR\s+A|DESPUES\s+DE)\s+(\d{1,2}/\d{1,2}/\d{4})", norm(description))
    if not match:
        # Search the original string because normalization removes punctuation.
        match = re.search(r"(?:mayor\s+a|posterior\s+a|despu[eé]s\s+de)\s+(\d{1,2}/\d{1,2}/\d{4})", description, re.I)
    if not match:
        return None
    try:
        return datetime.strptime(match.group(1), "%d/%m/%Y").date()
    except ValueError:
        return None


def infer_type(description: str) -> str:
    d = norm(description)
    if "ALFABETICO" in d or "ALFABETICA" in d:
        return "Texto alfabético"
    if "ALFANUMERICO" in d or "ALFANUMERICA" in d:
        return "Alfanumérico"
    if "NUMERICO" in d or "NUMERICA" in d:
        return "Numérico"
    if "DD MM AAAA" in d or "FORMATO" in d and ("FECHA" in d or "DIA" in d) or "FECHA" in d:
        return "Fecha"
    return "Sin validación de tipo"


def type_label(value: str) -> str:
    key = norm(value)
    mapping = {"TEXTO ALFABETICO": "Texto alfabético", "NUMERICO": "Numérico", "FECHA": "Fecha", "ALFANUMERICO": "Alfanumérico"}
    return mapping.get(key, "Sin validación de tipo")


def infer_required(description: str) -> bool:
    d = norm(description)
    return any(token in d for token in ("OBLIGATORIO", "OBLIGATORIA", "NO PUEDE ESTAR VACIO", "DEBE DILIGENCIAR"))


def infer_unique(description: str) -> bool:
    return "UNICO" in norm(description) or "UNICA" in norm(description)


def is_missing(value: object) -> bool:
    if value is None or pd.isna(value):
        return True
    return str(value).strip() == ""


def is_numeric(value: object) -> bool:
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return True
    text = str(value).strip().replace("$", "").replace(" ", "")
    # Permit standard decimal values and Colombian thousands separators.
    if re.fullmatch(r"[-+]?\d+(?:[.,]\d+)?", text):
        return True
    if re.fullmatch(r"[-+]?\d{1,3}(?:\.\d{3})+(?:,\d+)?", text):
        return True
    return False


def is_alpha(value: object) -> bool:
    text = str(value).strip()
    return bool(re.search(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]", text)) and not bool(re.search(r"\d", text))


def is_date(value: object) -> bool:
    if isinstance(value, (datetime, date, pd.Timestamp)):
        return True
    text = str(value).strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%m/%d/%Y"):
        try:
            datetime.strptime(text, fmt)
            return True
        except ValueError:
            pass
    return False


def check_value(value: object, rule: dict) -> list[str]:
    errors = []
    missing = is_missing(value)
    if missing:
        if rule["required"]:
            errors.append("Campo obligatorio sin diligenciar")
        return errors
    if norm(value) == "N A":
        return ["N/A no está permitido; use No Aplica solo si el campo lo admite"]
    if norm(value) == "NO APLICA":
        return [] if rule.get("allow_no_aplica", False) else ["No Aplica no está permitido para este campo"]
    kind = rule["type"]
    if kind == "Numérico" and not is_numeric(value):
        errors.append("Debe contener un valor numérico")
    elif kind == "Texto alfabético" and not is_alpha(value):
        errors.append("Debe contener texto alfabético (sin números)")
    elif kind == "Fecha" and not is_date(value):
        errors.append("Fecha inválida; use DD/MM/AAAA")
    elif kind == "Alfanumérico":
        # This is a descriptive type; any non-empty text or number is valid.
        pass
    max_date = rule.get("max_date")
    if kind == "Fecha" and max_date and not errors:
        parsed = value.date() if isinstance(value, (datetime, pd.Timestamp)) else value if isinstance(value, date) else None
        if parsed is None:
            for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
                try:
                    parsed = datetime.strptime(str(value).strip(), fmt).date()
                    break
                except ValueError:
                    pass
        if parsed and parsed > max_date:
            errors.append(f"La fecha no puede ser posterior a {max_date.strftime('%d/%m/%Y')}")
    allowed = rule.get("allowed_values", [])
    if allowed and not errors:
        candidate = norm(value)
        valid = any(candidate == norm(item) or candidate.endswith(norm(item)) for item in allowed)
        if not valid:
            errors.append("El valor no está en la lista de valores aceptados")
    return errors


def main() -> None:
    st.title("Malla de validación de Excel")
    st.write("Carga el libro con la información y las reglas. La app revisa la hoja elegida y genera un log descargable con las filas observadas.")
    data_file = st.file_uploader("Archivo Excel a auditar (XLSX)", type=["xlsx"])
    if not data_file:
        st.info("Carga un archivo para comenzar.")
        return

    try:
        data_sheets = read_sheets(data_file)
    except Exception as exc:
        st.error(f"No se pudo leer el archivo: {exc}")
        return

    names = list(data_sheets)
    rule_candidates = [n for n in names if "REGLA" in norm(n)]
    left, right = st.columns(2)
    with left:
        target_sheet = st.selectbox("Hoja que se va a auditar", names, index=next((i for i, n in enumerate(names) if norm(n) not in {"REGLAS", "REGLA"}), 0))
    with right:
        use_separate = st.checkbox("Las reglas están en otro archivo", value=False)

    if use_separate:
        rule_file = st.file_uploader("Archivo Excel de reglas (XLSX)", type=["xlsx"], key="rules_file")
        if not rule_file:
            st.info("Carga el archivo de reglas para continuar.")
            return
        try:
            rule_sheets = read_sheets(rule_file)
        except Exception as exc:
            st.error(f"No se pudo leer el archivo de reglas: {exc}")
            return
    else:
        rule_sheets = data_sheets
        if not rule_candidates:
            st.error("No encontré una hoja cuyo nombre contenga 'Regla'. Marca la opción para cargar un archivo de reglas separado.")
            return
    selected_rule_sheet = st.selectbox("Hoja de reglas", list(rule_sheets), index=(list(rule_sheets).index(rule_candidates[0]) if not use_separate else 0))

    raw_target = data_sheets[target_sheet]
    default_header = detect_header_row(raw_target)
    header_row = st.number_input("Fila de encabezados de la hoja a auditar", min_value=1, max_value=max(1, len(raw_target)), value=default_header + 1, step=1) - 1
    headers = [str(x).strip() if pd.notna(x) else "" for x in raw_target.iloc[int(header_row)].tolist()]
    valid_positions = [i for i, h in enumerate(headers) if h]
    if not valid_positions:
        st.error("No se encontraron encabezados en la fila seleccionada.")
        return
    if len(set(headers[i] for i in valid_positions)) != len(valid_positions):
        st.warning("Hay encabezados repetidos. Se distinguirán por su posición en Excel.")
    display_headers = [f"{headers[i]} (columna {i + 1})" for i in valid_positions]
    raw_rules = rule_sheets[selected_rule_sheet]
    rules, _ = parse_rules(raw_rules)
    if not rules:
        st.error("No pude identificar reglas. Use columnas COLUMNA y DESCRIPCIÓN, o CAMPO DE REGLA, TIPO_CAMPO, OBLIGATORIO, NO APLICA y VALORES ACEPTADOS.")
        return

    st.subheader("Revisión de reglas y campos")
    st.caption("Confirma el emparejamiento campo a columna y revisa las opciones cargadas desde las reglas. Puedes ajustar tipo, obligatoriedad, ‘No Aplica’ y valores aceptados antes de validar.")
    configured = []
    for idx, rule in enumerate(rules):
        label = f"{rule['field']} — {rule['description'].splitlines()[0][:130]}"
        keybase = f"rule_{idx}"
        rule_norm = norm(rule["field"])
        best = max(valid_positions, key=lambda p: SequenceMatcher(None, rule_norm, norm(headers[p])).ratio())
        exact = [p for p in valid_positions if norm(headers[p]) == rule_norm]
        default_pos = exact[0] if exact else best
        options = ["(No validar)"] + display_headers
        default_opt = valid_positions.index(default_pos) + 1
        with st.expander(label, expanded=False):
            chosen = st.selectbox("Columna de datos", options, index=default_opt, key=keybase + "_col")
            c1, c2, c3 = st.columns(3)
            with c1:
                types = ["Sin validación de tipo", "Texto alfabético", "Numérico", "Fecha", "Alfanumérico"]
                suggested_type = type_label(rule.get("declared_type", "")) if rule.get("structured") else infer_type(rule["description"])
                kind = st.selectbox("Tipo", types, index=types.index(suggested_type), key=keybase + "_type")
            with c2:
                required = st.checkbox("Campo obligatorio", value=rule.get("required_default", infer_required(rule["description"])), key=keybase + "_required")
            with c3:
                unique = st.checkbox("Debe ser único", value=infer_unique(rule["description"]), key=keybase + "_unique")
            allow_no_aplica = st.checkbox("Aceptar exactamente ‘No Aplica’", value=rule.get("no_aplica_default", False), key=keybase + "_noaplica")
            allowed_default = rule.get("allowed_default", "")
            allowed_text = st.text_area("Valores aceptados (uno por línea; vacío = sin catálogo)", value=allowed_default, key=keybase + "_allowed")
            if rule.get("max_date"):
                st.caption(f"Límite máximo de fecha detectado: {rule['max_date'].strftime('%d/%m/%Y')}")
            st.caption(rule["description"])
        if chosen != "(No validar)":
            col_index = valid_positions[options.index(chosen) - 1]
            configured.append({**rule, "column_index": col_index, "column": headers[col_index], "type": kind, "required": required, "unique": unique,
                               "allow_no_aplica": allow_no_aplica, "allowed_values": [x.strip() for x in allowed_text.splitlines() if x.strip()]})

    if not configured:
        st.warning("Selecciona al menos una columna para validar.")
        return
    if not st.button("Ejecutar validación", type="primary"):
        return

    data = raw_target.iloc[int(header_row) + 1 :, valid_positions].copy()
    data.columns = [headers[i] for i in valid_positions]
    # Preserve original Excel row number for each record and discard completely empty rows.
    excel_rows = [int(header_row) + 2 + i for i in range(len(data))]
    issues = []
    counts_by_rule = {r["field"]: 0 for r in configured}
    for row_pos, (_, row) in enumerate(data.iterrows()):
        if all(is_missing(x) for x in row.tolist()):
            continue
        excel_row = excel_rows[row_pos]
        for rule in configured:
            value = row.iloc[valid_positions.index(rule["column_index"])]
            failures = check_value(value, rule)
            if rule["unique"] and not is_missing(value):
                series = data.iloc[:, valid_positions.index(rule["column_index"])]
                matches = series.astype(str).str.strip().str.casefold() == str(value).strip().casefold()
                if int(matches.sum()) > 1:
                    failures.append("El valor está duplicado; debe ser único")
            for message in failures:
                counts_by_rule[rule["field"]] += 1
                issues.append({"Fila Excel": excel_row, "Campo": rule["field"], "Columna": rule["column"], "Valor encontrado": str(value), "Observación": message})

    total_data = sum(not all(is_missing(x) for x in row) for row in data.itertuples(index=False, name=None))
    issue_df = pd.DataFrame(issues, columns=["Fila Excel", "Campo", "Columna", "Valor encontrado", "Observación"])
    bad_rows = issue_df["Fila Excel"].nunique() if not issue_df.empty else 0
    summary = pd.DataFrame([
        {"Métrica": "Hoja auditada", "Resultado": target_sheet},
        {"Métrica": "Filas de datos revisadas", "Resultado": total_data},
        {"Métrica": "Filas con observaciones", "Resultado": bad_rows},
        {"Métrica": "Observaciones", "Resultado": len(issue_df)},
        {"Métrica": "Resultado", "Resultado": "Con observaciones" if len(issue_df) else "Sin errores detectados"},
    ])
    st.subheader("Resultado")
    m1, m2, m3 = st.columns(3)
    m1.metric("Filas revisadas", total_data)
    m2.metric("Filas con observaciones", bad_rows)
    m3.metric("Observaciones", len(issue_df))
    if issue_df.empty:
        st.success("No se detectaron errores con las reglas seleccionadas.")
    else:
        st.dataframe(issue_df, use_container_width=True, hide_index=True)
    st.download_button("Descargar log Excel", data=make_log(summary, issue_df, configured), file_name="log_validacion.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("Descargar log CSV", data=issue_df.to_csv(index=False).encode("utf-8-sig"), file_name="log_validacion.csv", mime="text/csv")


def make_log(summary: pd.DataFrame, issues: pd.DataFrame, rules: list[dict]) -> bytes:
    configured_rules = pd.DataFrame([{"Campo de regla": r["field"], "Columna auditada": r["column"], "Tipo aplicado": r["type"], "Obligatorio": r["required"], "Único": r["unique"], "Acepta No Aplica": r.get("allow_no_aplica", False), "Valores aceptados": " | ".join(r.get("allowed_values", [])), "Fecha máxima": r["max_date"].strftime("%d/%m/%Y") if r.get("max_date") else "", "Descripción": r["description"]} for r in rules])
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Resumen", index=False)
        issues.to_excel(writer, sheet_name="Observaciones", index=False)
        configured_rules.to_excel(writer, sheet_name="Reglas aplicadas", index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for column_cells in ws.columns:
                width = min(70, max(12, max((len(str(c.value or "")) for c in column_cells), default=10) + 2))
                ws.column_dimensions[column_cells[0].column_letter].width = width
    return buffer.getvalue()


if __name__ == "__main__":
    main()
