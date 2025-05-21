import flet as ft
import os
import re
import json
import extract_operations

GREY_C = "#2A2D33"
HEADER_RE = re.compile(r'^[A-Z]+(?:,[A-Z]+)*$')


def _on_header_focus(e: ft.ControlEvent):
    e.control.border_color = ft.Colors.BLUE_800
    e.page.update()


def make_sheet_info(filename: str):
    txt_headers = ft.TextField(
        value="F,J,N,R,V,Z,AD,AH,AL,AP,AT,AX,BB,BF,BJ",
        expand=True,
        border_color=ft.Colors.BLUE_800,
        focused_border_color=ft.Colors.BLUE,
        on_focus=lambda e: _on_header_focus(e)
    )
    txt_first = ft.TextField(
        value="3",
        width=60,
        border_color=ft.Colors.BLUE_800,
        focused_border_color=ft.Colors.BLUE,
        on_focus=lambda e: _on_header_focus(e)
    )
    txt_last = ft.TextField(
        value="39",
        width=60,
        border_color=ft.Colors.BLUE_800,
        focused_border_color=ft.Colors.BLUE,
        on_focus=lambda e: _on_header_focus(e)
    )
    c = ft.Container(
        ft.Column([
            ft.Text(f"{filename}:"),
            ft.Row([
                ft.Text("Cabeçalhos:"),
                txt_headers,
                ft.Text("Linha primeiro Aluno:"),
                txt_first,
                ft.Text("Linha último Aluno:"),
                txt_last
            ], spacing=10)
        ], spacing=15),
        padding=15,
        border=ft.border.all(1, GREY_C),
        border_radius=5,
        expand=True
    )
    c.txt_fields = [txt_headers, txt_first, txt_last]
    return c


def main(page: ft.Page):
    page.title = "Extrator JSON de Alunos em Recuperação"
    page.theme_mode = "dark"
    page.window.width = 790
    page.window.height = 265
    page.padding = 20

    txt_path = ft.TextField(
        label="Selecionar Planilhas:",
        expand=True,
        read_only=True,
        border_color=ft.Colors.BLUE_800,
        focused_border_color=ft.Colors.BLUE
    )

    def on_file_result(e: ft.FilePickerResultEvent):
        if e.files:
            txt_path.value = ", ".join(f.path for f in e.files)
            txt_path.border_color = ft.Colors.BLUE_800
            page.update()

    file_picker = ft.FilePicker(on_result=on_file_result)
    page.overlay.append(file_picker)

    btn_upload = ft.ElevatedButton(
        text="Upload",
        width=200,
        height=50,
        on_click=lambda e: file_picker.pick_files(allow_multiple=True)
    )
    btn_carregar = ft.ElevatedButton(text="Carregar", expand=True, height=50)
    btn_load = ft.ElevatedButton(text="Carregar dados em JSON", expand=True, height=40)

    initial_page = ft.Column([
        ft.Row([ft.Text(
            "Extrator JSON de Alunos em Recuperação",
            style="headlineMedium",
            color=ft.Colors.BLUE
        )], alignment=ft.MainAxisAlignment.CENTER),
        ft.Divider(thickness=1),
        ft.Row([txt_path, btn_upload], spacing=10),
        ft.Row([btn_carregar], spacing=10)
    ], expand=True)

    afterFileSelecion = ft.Column([], expand=True, visible=False)
    sheet_containers: list[ft.Container] = []

    def on_load(e):
        # validação de campos
        for sheet in sheet_containers:
            for tf in sheet.txt_fields:
                if not tf.value.strip():
                    tf.border_color = ft.Colors.RED
                    page.open(ft.SnackBar(
                        ft.Text("Por favor, preencha os dados!"),
                        bgcolor=ft.Colors.RED_100
                    ))
                    page.update()
                    return
            headers_tf, first_tf, last_tf = sheet.txt_fields
            if not HEADER_RE.match(headers_tf.value.strip()):
                headers_tf.border_color = ft.Colors.YELLOW
                page.open(ft.SnackBar(
                    ft.Text("Cabeçalhos inválidos! Use somente letras maiúsculas separadas por vírgula!"),
                    bgcolor=ft.Colors.YELLOW_100
                ))
                page.update()
                return
            val1 = first_tf.value.strip()
            if not val1.isdigit() or int(val1) <= 0:
                first_tf.border_color = ft.Colors.YELLOW
                page.open(ft.SnackBar(
                    ft.Text("Linha primeiro Aluno deve ser inteiro > 0!"),
                    bgcolor=ft.Colors.YELLOW_100
                ))
                page.update()
                return
            val2 = last_tf.value.strip()
            if not val2.isdigit() or int(val2) <= int(val1):
                last_tf.border_color = ft.Colors.YELLOW
                page.open(ft.SnackBar(
                    ft.Text("Linha último Aluno deve ser inteiro > primeiro!"),
                    bgcolor=ft.Colors.YELLOW_100
                ))
                page.update()
                return
        # extrai e mescla JSONs
        paths = [p.strip() for p in txt_path.value.split(",")]
        jsons = []
        for idx, sheet in enumerate(sheet_containers):
            headers = sheet.txt_fields[0].value
            firstRow = int(sheet.txt_fields[1].value)
            lastRow = int(sheet.txt_fields[2].value)
            jsons.append(
                extract_operations.extract_json(
                    paths[idx], headers, firstRow, lastRow
                )
            )
        merged = extract_operations.merge_jsons(jsons)
        # exporta para arquivo
        output_file = "AlunosEmRecuperacao.json"
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(merged, f, ensure_ascii=False, indent=2)
        page.open(ft.SnackBar(
            ft.Text(f"Exportado '{output_file}' com sucesso!"),
            bgcolor=ft.Colors.GREEN_100
        ))
        page.update()

    btn_load.on_click = on_load

    def switch_page(e):
        if not txt_path.value:
            txt_path.border_color = ft.Colors.RED
            page.open(ft.SnackBar(
                ft.Text("Por favor, selecione as planilhas!"),
                bgcolor=ft.Colors.RED_100
            ))
            page.update()
            return
        paths = [p.strip() for p in txt_path.value.split(",")]
        if any(not p.lower().endswith(".xlsx") for p in paths):
            txt_path.border_color = ft.Colors.YELLOW
            page.open(ft.SnackBar(
                ft.Text("Selecione as planilhas no formato .xlsx!"),
                bgcolor=ft.Colors.YELLOW_100
            ))
            page.update()
            return
        sheet_containers.clear()
        controls = [
            ft.Row([ft.Text(
                "Extrator JSON de Alunos em Recuperação",
                style="headlineMedium",
                color=ft.Colors.BLUE
            )], alignment=ft.MainAxisAlignment.CENTER),
            ft.Divider(thickness=1),
            *[
                sheet_containers.append(
                    make_sheet_info(os.path.basename(p))
                ) or sheet_containers[-1]
                for p in paths
            ],
            ft.Row([btn_load], spacing=10)
        ]
        afterFileSelecion.controls = controls
        initial_page.visible = False
        afterFileSelecion.visible = True
        page.window.maximized = True
        page.update()

    btn_carregar.on_click = switch_page
    page.add(initial_page, afterFileSelecion)

if __name__ == "__main__":
    ft.app(target=main)
