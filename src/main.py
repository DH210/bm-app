import csv
import json
import os
import sys
from collections import defaultdict
import PySimpleGUIQt as sg  # Backend Qt (compatível Win/Linux)

APP_TITLE = "CBMERJ Cursos - Cadastro e Análise"
COURSES_FILE = "courses.json"

BASE_HEADERS = [
    "ID", "Curso", "Ano", "Vagas", "Inscritos",
    "Aptos", "Matriculados", "Desligados", "Motivos", "Concluintes"
]
DISPLAY_HEADERS = BASE_HEADERS + ["% Conclusão", "% Evasão"]

def resource_path(rel_path: str) -> str:
    if hasattr(sys, "_MEIPASS"):  # PyInstaller
        return os.path.join(sys._MEIPASS, rel_path)
    return os.path.join(os.path.abspath("."), rel_path)

def load_courses():
    path = resource_path(COURSES_FILE)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return sorted(set([str(x).strip() for x in data if str(x).strip()]))
        except Exception:
            pass
    # Lista inicial (edite depois em Dados > Gerenciar cursos)
    seed = [
        "Curso A", "Curso B", "Curso C"
    ]
    save_courses(seed)
    return seed

def save_courses(courses):
    path = resource_path(COURSES_FILE)
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(sorted(set([str(c).strip() for c in courses if str(c).strip()])), f, ensure_ascii=False, indent=2)
    except Exception as e:
        sg.popup_error(f"Erro ao salvar lista de cursos: {e}", keep_on_top=True)

def to_int(val, default=0):
    try:
        return int(str(val).strip())
    except Exception:
        return default

def format_pct(num, den):
    if den and den > 0:
        return f"{(num/den)*100:.1f}%"
    return "0.0%"

def compute_display_row(rec):
    concl = to_int(rec["Concluintes"])
    matr = to_int(rec["Matriculados"])
    desl  = to_int(rec["Desligados"])
    pct_conc = format_pct(concl, matr)
    pct_evas = format_pct(desl, matr)
    base = [
        rec["ID"], rec["Curso"], rec["Ano"], rec["Vagas"], rec["Inscritos"],
        rec["Aptos"], rec["Matriculados"], rec["Desligados"], rec["Motivos"], rec["Concluintes"]
    ]
    return base + [pct_conc, pct_evas]

def validate_record(rec):
    vagas = to_int(rec["Vagas"])
    insc  = to_int(rec["Inscritos"])
    aptos = to_int(rec["Aptos"])
    matr  = to_int(rec["Matriculados"])
    desl  = to_int(rec["Desligados"])
    concl = to_int(rec["Concluintes"])

    msgs = []
    if not str(rec["Curso"]).strip():
        msgs.append("Curso é obrigatório.")
    if to_int(rec["Ano"]) <= 1900:
        msgs.append("Ano inválido (use 4 dígitos).")
    for k, v in [("Vagas", vagas), ("Inscritos", insc), ("Aptos", aptos),
                 ("Matriculados", matr), ("Desligados", desl), ("Concluintes", concl)]:
        if v < 0:
            msgs.append(f"{k} não pode ser negativo.")

    if aptos > insc:
        msgs.append("Aptos não pode ser maior que Inscritos.")
    if matr > aptos:
        msgs.append("Matriculados não pode ser maior que Aptos.")
    if desl > matr:
        msgs.append("Desligados não pode ser maior que Matriculados.")
    if concl > matr:
        msgs.append("Concluintes não pode ser maior que Matriculados.")
    if (desl + concl) > matr:
        msgs.append("Desligados + Concluintes não pode exceder Matriculados.")

    return (len(msgs) == 0), "\n".join(msgs)

def new_record(next_id, values):
    return {
        "ID": next_id,
        "Curso": values.get("-CURSO-", "").strip(),
        "Ano": to_int(values.get("-ANO-", "")),
        "Vagas": to_int(values.get("-VAGAS-", "")),
        "Inscritos": to_int(values.get("-INSCR-", "")),
        "Aptos": to_int(values.get("-APTOS-", "")),
        "Matriculados": to_int(values.get("-MATR-", "")),
        "Desligados": to_int(values.get("-DESL-", "")),
        "Motivos": values.get("-MOT-", "").strip(),
        "Concluintes": to_int(values.get("-CONC-", "")),
    }

def rec_to_base_row(rec):
    return [rec[h] for h in BASE_HEADERS]

def base_row_to_rec(row):
    return {
        "ID": to_int(row[0]),
        "Curso": str(row[1]),
        "Ano": to_int(row[2]),
        "Vagas": to_int(row[3]),
        "Inscritos": to_int(row[4]),
        "Aptos": to_int(row[5]),
        "Matriculados": to_int(row[6]),
        "Desligados": to_int(row[7]),
        "Motivos": str(row[8]),
        "Concluintes": to_int(row[9]),
    }

def save_csv(path, records):
    with open(path, "w", newline="", encoding="utf-8") as f:
        wr = csv.writer(f)
        wr.writerow(BASE_HEADERS)
        for rec in records:
            wr.writerow(rec_to_base_row(rec))

def load_csv(path):
    out = []
    with open(path, "r", newline="", encoding="utf-8") as f:
        rd = csv.reader(f)
        headers = next(rd, None)
        for row in rd:
            row = (row + [""] * len(BASE_HEADERS))[:len(BASE_HEADERS)]
            out.append(base_row_to_rec(row))
    return out

def filter_records(records, courses_sel, ano_min, ano_max, motivo_txt):
    def ok(rec):
        if courses_sel and rec["Curso"] not in courses_sel:
            return False
        ano = to_int(rec["Ano"])
        if ano_min and ano < ano_min:
            return False
        if ano_max and ano > ano_max:
            return False
        if motivo_txt and motivo_txt.lower() not in rec["Motivos"].lower():
            return False
        return True
    return [r for r in records if ok(r)]

def aggregate(records, group_by="Curso"):
    agg = defaultdict(lambda: {
        "Vagas": 0, "Inscritos": 0, "Aptos": 0, "Matriculados": 0, "Desligados": 0, "Concluintes": 0
    })
    for r in records:
        key = r[group_by]
        a = agg[key]
        a["Vagas"]        += to_int(r["Vagas"])
        a["Inscritos"]    += to_int(r["Inscritos"])
        a["Aptos"]        += to_int(r["Aptos"])
        a["Matriculados"] += to_int(r["Matriculados"])
        a["Desligados"]   += to_int(r["Desligados"])
        a["Concluintes"]  += to_int(r["Concluintes"])

    headers = [group_by, "Vagas", "Inscritos", "Aptos", "Matriculados",
               "Desligados", "Concluintes", "% Conclusão", "% Evasão"]
    rows = []
    for key, a in sorted(agg.items(), key=lambda kv: str(kv[0])):
        pct_conc = format_pct(a["Concluintes"], a["Matriculados"])
        pct_evas = format_pct(a["Desligados"], a["Matriculados"])
        rows.append([
            key, a["Vagas"], a["Inscritos"], a["Aptos"], a["Matriculados"],
            a["Desligados"], a["Concluintes"], pct_conc, pct_evas
        ])
    return headers, rows

def manage_courses_dialog(courses):
    layout = [
        [sg.Text("Cursos (um por linha):")],
        [sg.Multiline("\n".join(courses), key="-ML-", size=(50, 15))],
        [sg.Button("Salvar"), sg.Button("Cancelar")]
    ]
    w = sg.Window("Gerenciar Cursos", layout, modal=True)
    new_list = courses[:]
    while True:
        ev, vals = w.read()
        if ev in (sg.WIN_CLOSED, "Cancelar"):
            break
        if ev == "Salvar":
            lines = [l.strip() for l in vals["-ML-"].splitlines()]
            new_list = [l for l in lines if l]
            save_courses(new_list)
            sg.popup_ok("Lista de cursos atualizada.", keep_on_top=True)
            break
    w.close()
    return new_list

def make_layout(courses):
    menu_def = [
        ["Arquivo", ["Novo", "Abrir CSV", "Salvar CSV", "Exportar filtrados", "Exportar agregado", "---", "Sair"]],
        ["Dados", ["Gerenciar cursos"]],
        ["Ajuda", ["Como usar", "Sobre"]]
    ]

    cadastro_tab = [
        [sg.Text("Curso"), sg.Combo(courses, key="-CURSO-", readonly=True, size=(40, 1), expand_x=True)],
        [sg.Text("Ano", size=(12,1)), sg.Input(key="-ANO-", size=(10,1)),
         sg.Text("Vagas", size=(10,1)), sg.Input(key="-VAGAS-", size=(10,1)),
         sg.Text("Inscritos", size=(10,1)), sg.Input(key="-INSCR-", size=(10,1))],
        [sg.Text("Aptos", size=(12,1)), sg.Input(key="-APTOS-", size=(10,1)),
         sg.Text("Matriculados", size=(10,1)), sg.Input(key="-MATR-", size=(10,1)),
         sg.Text("Desligados", size=(10,1)), sg.Input(key="-DESL-", size=(10,1))],
        [sg.Text("Concluintes", size=(12,1)), sg.Input(key="-CONC-", size=(10,1))],
        [sg.Text("Motivos (texto livre)"),],
        [sg.Multiline(key="-MOT-", size=(70,5))],
        [sg.Button("Adicionar", key="-ADD-"), sg.Button("Limpar", key="-CLR-")],
    ]

    tabela_tab = [
        [sg.Table(values=[],
                  headings=DISPLAY_HEADERS,
                  key="-TBL-",
                  expand_x=True, expand_y=True,
                  enable_events=True,
                  alternating_row_color="#f6f8fa",
                  select_mode=sg.TABLE_SELECT_MODE_EXTENDED,
                  num_rows=15,
                  auto_size_columns=True)],
        [sg.Button("Remover seleção", key="-DEL-")]
    ]

    filtros_tab = [
        [sg.Text("Cursos"), sg.Listbox(values=courses, key="-FCURSOS-", size=(40,8), select_mode=sg.LISTBOX_SELECT_MODE_MULTIPLE, expand_y=True)],
        [sg.Text("Ano mín"), sg.Input(key="-FANOMIN-", size=(8,1)),
         sg.Text("Ano máx"), sg.Input(key="-FANOMAX-", size=(8,1)),
         sg.Text("Motivo contém"), sg.Input(key="-FMOT-", size=(20,1))],
        [sg.Button("Aplicar filtro", key="-APLFILT-"), sg.Button("Limpar filtro", key="-CLRFILT-")],
        [sg.Text("Total registros filtrados:"), sg.Text("0", key="-FCOUNT-")]
    ]

    analise_tab = [
        [sg.Text("Agrupar por:"), sg.Combo(["Curso", "Ano"], default_value="Curso", readonly=True, key="-GROUPBY-"),
         sg.Button("Recalcular", key="-RECALC-")],
        [sg.Table(values=[],
                  headings=["Grupo","Vagas","Inscritos","Aptos","Matriculados","Desligados","Concluintes","% Conclusão","% Evasão"],
                  key="-TBLAGG-",
                  expand_x=True, expand_y=True,
                  alternating_row_color="#f6f8fa",
                  num_rows=12,
                  auto_size_columns=True)],
    ]

    log_tab = [
        [sg.Multiline(size=(80,10), key="-LOG-", autoscroll=True, reroute_stdout=False, reroute_stderr=False)]
    ]

    layout = [
        [sg.Menu(menu_def)],
        [sg.TabGroup([[
            sg.Tab("Cadastro", cadastro_tab),
            sg.Tab("Tabela", tabela_tab),
            sg.Tab("Filtros", filtros_tab),
            sg.Tab("Análise", analise_tab),
            sg.Tab("Log", log_tab),
        ]], expand_x=True, expand_y=True)],
        [sg.StatusBar("Pronto", key="-STATUS-", size=(80,1))]
    ]
    return layout

def log(window, msg):
    window["-LOG-"].print(msg)
    window["-STATUS-"].update(msg)

def popup_sobre():
    sg.popup_ok(
        "CBMERJ Cursos - Cadastro e Análise\n\n"
        "• Cadastro por curso/ano com consistência básica\n"
        "• Filtros por curso/ano/motivo\n"
        "• Análise agregada por Curso ou Ano (taxas)\n"
        "• CSV: abrir, salvar e exportar relatórios",
        title="Sobre", keep_on_top=True
    )

def popup_help():
    sg.popup_ok(
        "Como usar (resumo):\n\n"
        "1) Dados > Gerenciar cursos: ajuste a lista de cursos (salva em courses.json).\n"
        "2) Aba Cadastro: preencha os campos e clique 'Adicionar'.\n"
        "3) Aba Tabela: visualize, selecione e remova linhas se necessário.\n"
        "4) Aba Filtros: selecione cursos/ano e clique 'Aplicar filtro'.\n"
        "   • 'Exportar filtrados' (menu Arquivo) salva só o que está filtrado.\n"
        "5) Aba Análise: escolha 'Curso' ou 'Ano' e clique 'Recalcular'.\n"
        "   • 'Exportar agregado' (menu Arquivo) salva a tabela agregada.\n"
        "6) Arquivo > Abrir/Salvar CSV: persiste seus dados base.",
        title="Como usar", keep_on_top=True
    )

def choose_save_file(title, default_name, file_pattern):
    return sg.popup_get_file(title, save_as=True, default_extension=".csv",
                             initial_filename=default_name,
                             file_types=((file_pattern, "*.csv"),),
                             no_window=True)

def choose_open_file(title):
    return sg.popup_get_file(title, file_types=(("CSV", "*.csv"),), no_window=True)

def export_filtered(window, filtered_rows):
    if not filtered_rows:
        sg.popup_ok("Não há dados filtrados para exportar.", keep_on_top=True)
        return
    path = choose_save_file("Exportar CSV (filtrados)", "dados_filtrados.csv", "CSV")
    if not path:
        return
    try:
        with open(path, "w", newline="", encoding="utf-8") as f:
            wr = csv.writer(f)
            wr.writerow(DISPLAY_HEADERS)
            for row in filtered_rows:
                wr.writerow(row)
        log(window, f"Exportado filtrado (com colunas calculadas) em: {path}")
    except Exception as e:
        sg.popup_error(f"Erro ao exportar: {e}", keep_on_top=True)

def export_agg(window, agg_headers, agg_rows):
    if not agg_rows:
        sg.popup_ok("Não há dados agregados para exportar.", keep_on_top=True)
        return
    path = choose_save_file("Exportar CSV (agregado)", "relatorio_agregado.csv", "CSV")
    if not path:
        return
    try:
        with open(path, "w", newline="", encoding="utf-8") as f:
            wr = csv.writer(f)
            wr.writerow(agg_headers)
            for row in agg_rows:
                wr.writerow(row)
        log(window, f"Relatório agregado exportado em: {path}")
    except Exception as e:
        sg.popup_error(f"Erro ao exportar: {e}", keep_on_top=True)

def main():
    sg.theme("LightGrey1")
    courses = load_courses()

    window = sg.Window(APP_TITLE, make_layout(courses), resizable=True, finalize=True)

    records = []
    filtered = []
    next_id = 1
    current_csv = ""

    def refresh_table(view=None):
        nonlocal filtered
        view_rows = view if view is not None else [compute_display_row(r) for r in records]
        filtered = view_rows
        window["-TBL-"].update(values=view_rows)
        window["-FCOUNT-"].update(str(len(view_rows)))

    def apply_filter():
        ano_min = to_int(window["-FANOMIN-"].get()) if window["-FANOMIN-"].get().strip() else None
        ano_max = to_int(window["-FANOMAX-"].get()) if window["-FANOMAX-"].get().strip() else None
        motivo  = window["-FMOT-"].get().strip()
        cursos_sel = window["-FCURSOS-"].get() or []
        recs = filter_records(records, cursos_sel, ano_min, ano_max, motivo)
        view = [compute_display_row(r) for r in recs]
        refresh_table(view)

    def recalc_agg():
        group_by = window["-GROUPBY-"].get() or "Curso"
        headers, rows = aggregate(records, group_by=group_by)
        window["-TBLAGG-"].update(values=rows, headings=headers)

    refresh_table()
    recalc_agg()

    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, "Sair"):
            break

        if event == "Sobre":
            popup_sobre()
        if event == "Como usar":
            popup_help()

        if event == "Gerenciar cursos":
            new_list = manage_courses_dialog(courses)
            courses = new_list
            window["-CURSO-"].update(values=courses)
            window["-FCURSOS-"].update(values=courses)

        if event in ("Novo",):
            records = []
            filtered = []
            next_id = 1
            current_csv = ""
            refresh_table()
            recalc_agg()
            log(window, "Novo dataset iniciado.")

        if event in ("Abrir CSV",):
            path = choose_open_file("Abrir CSV")
            if path:
                try:
                    records = load_csv(path)
                    current_csv = path
                    try:
                        max_id = max(r["ID"] for r in records) if records else 0
                        next_id = max_id + 1
                    except Exception:
                        next_id = 1
                    refresh_table()
                    recalc_agg()
                    log(window, f"CSV carregado: {path}")
                except Exception as e:
                    sg.popup_error(f"Erro ao carregar CSV: {e}", keep_on_top=True)

        if event in ("Salvar CSV",):
            path = current_csv or choose_save_file("Salvar CSV", "dados_cursos.csv", "CSV")
            if path:
                try:
                    save_csv(path, records)
                    current_csv = path
                    log(window, f"CSV salvo em: {path}")
                except Exception as e:
                    sg.popup_error(f"Erro ao salvar CSV: {e}", keep_on_top=True)

        if event == "Exportar filtrados":
            export_filtered(window, filtered)

        if event == "Exportar agregado":
            group_by = window["-GROUPBY-"].get() or "Curso"
            headers, rows = aggregate(records, group_by=group_by)
            export_agg(window, headers, rows)

        if event == "-ADD-":
            rec = new_record(next_id, values)
            ok, msg = validate_record(rec)
            if not ok:
                sg.popup_error(f"Verifique os dados:\n\n{msg}", keep_on_top=True)
            else:
                records.append(rec)
                next_id += 1
                refresh_table()
                recalc_agg()
                log(window, f"Adicionado: {rec}")

        if event == "-CLR-":
            for k in ("-CURSO-","-ANO-","-VAGAS-","-INSCR-","-APTOS-","-MATR-","-DESL-","-CONC-","-MOT-"):
                window[k].update("")
            log(window, "Formulário limpo.")

        if event == "-DEL-":
            sel = window["-TBL-"].get_selected_rows()
            if not sel:
                sg.popup_ok("Selecione linhas para remover.", keep_on_top=True)
            else:
                try:
                    ids_to_remove = []
                    for idx in sel:
                        row = window["-TBL-"].Values[idx]
                        rec_id = to_int(row[0])
                        ids_to_remove.append(rec_id)
                    before = len(records)
                    records = [r for r in records if r["ID"] not in ids_to_remove]
                    after = len(records)
                    log(window, f"Removidas {before-after} linha(s).")
                    refresh_table()
                    recalc_agg()
                except Exception as e:
                    sg.popup_error(f"Erro ao remover: {e}", keep_on_top=True)

        if event == "-APLFILT-":
            apply_filter()

        if event == "-CLRFILT-":
            window["-FCURSOS-"].update(set_to_index=[])
            window["-FANOMIN-"].update("")
            window["-FANOMAX-"].update("")
            window["-FMOT-"].update("")
            refresh_table()

        if event == "-RECALC-":
            recalc_agg()

    window.close()

if __name__ == "__main__":
    main()
