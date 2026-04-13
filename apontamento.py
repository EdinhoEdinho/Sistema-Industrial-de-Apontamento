import tkinter as tk

from tkinter import ttk, messagebox

try:

    import pyodbc

except Exception:

    pyodbc = None

from datetime import datetime


conn_str = (

    'DRIVER={ODBC Driver 17 for SQL Server};'

    'SERVER=;' #preencher conforme

    'DATABASE=;' #preencher conforme

    'UID=;' #preencher conforme

    'PWD=;' #preencher conforme

    'Timeout=5;'

)


conn = None

db_online = False

db_error = ""

if pyodbc is None:

    db_error = "pyodbc não está instalado no ambiente."

    print(f"Sem conexão com banco. Interface iniciará em modo offline. Erro: {db_error}")

else:

    try:

        conn = pyodbc.connect(conn_str)

        db_online = True

        print("Conectado com sucesso!")

    except Exception as e:

        db_error = str(e)

        conn = None

        print(f"Sem conexão com banco. Interface iniciará em modo offline. Erro: {e}")


min_esticagem = 0

max_esticagem = 0

min_pos_esticagem = 0

max_pos_esticagem = 0


def carregar_parametros():

    global min_esticagem, max_esticagem, min_pos_esticagem, max_pos_esticagem

    try:

        cursor = conn.cursor()

        cursor.execute("SELECT TOP 1 minimo, maximo FROM parametro_esticagem ORDER BY id DESC")

        row = cursor.fetchone()

        if row:

            min_esticagem, max_esticagem = row

        else:

            min_esticagem, max_esticagem = 0, 0


        cursor.execute("SELECT TOP 1 minimo, maximo FROM parametro_pos_esticagem ORDER BY id DESC")

        row = cursor.fetchone()

        if row:

            min_pos_esticagem, max_pos_esticagem = row

        else:

            min_pos_esticagem, max_pos_esticagem = 0, 0


        cursor.close()

    except Exception as e:

        messagebox.showerror("Erro", f"Erro ao carregar parâmetros: {e}")

        min_esticagem, max_esticagem = 0, 0

        min_pos_esticagem, max_pos_esticagem = 0, 0


def validar_entry(event):

    widget = event.widget

    valor = widget.get().replace(',', '.').strip()


    if valor == "":

        widget.delete(0, tk.END)

        widget.configure(bg="white")

        return


    try:

        num_val = float(valor)

        # --- Formata no máximo 2 casas decimais ---

        s = f"{num_val:.10f}".rstrip('0').rstrip('.')

        if '.' in s:

            partes = s.split('.')

            s = f"{partes[0]}.{partes[1][:2]}"

        valor_formatado = s


        widget.delete(0, tk.END)

        widget.insert(0, valor_formatado)


    except ValueError:

        messagebox.showwarning("Aviso", "Valor inválido. Digite um número.")

        widget.focus_set()

        widget.configure(bg="red")

        return


    # --- Esticagem ---

    if widget in [e for pair in entries_tensao for e in pair]:

        status = ""

        algum_preenchido = False

        status_ng = False


        for entry_x, entry_y in entries_tensao:

            for entry in [entry_x, entry_y]:

                val = entry.get().replace(',', '.').strip()

                if val == "":

                    continue

                algum_preenchido = True

                try:

                    num_val = float(val)

                    s = f"{num_val:.10f}".rstrip('0').rstrip('.')

                    if '.' in s:

                        partes = s.split('.')

                        s = f"{partes[0]}.{partes[1][:2]}"

                    entry.delete(0, tk.END)

                    entry.insert(0, s)


                    if num_val < float(min_esticagem) or num_val > float(max_esticagem):

                        status = "NG"

                        status_ng = True

                        entry.configure(bg="red")  # causador vermelho

                    else:

                        entry.configure(bg="white")

                        if not status_ng:

                            status = "OK"

                except ValueError:

                    status = "NG"

                    status_ng = True

                    entry.configure(bg="red")


        entry_status.config(state="normal")

        entry_status.delete(0, tk.END)

        entry_status.insert(0, status)


        if status == "NG":

            entry_status.config(state="readonly", readonlybackground="red", fg="white")

        elif status == "OK":

            entry_status.config(state="normal", bg="white", fg="black")

        else:

            entry_status.config(state="normal", bg="white", fg="black")


    # --- Pós-Esticagem ---

    elif widget in [e for pair in entries_tensao_3d for e in pair]:

        status_pos = ""

        algum_preenchido = False

        status_ng = False


        for entry_x, entry_y in entries_tensao_3d:

            for entry in [entry_x, entry_y]:

                val = entry.get().replace(',', '.').strip()

                if val == "":

                    continue

                algum_preenchido = True

                try:

                    num_val = float(val)

                    s = f"{num_val:.10f}".rstrip('0').rstrip('.')

                    if '.' in s:

                        partes = s.split('.')

                        s = f"{partes[0]}.{partes[1][:2]}"

                    entry.delete(0, tk.END)

                    entry.insert(0, s)


                    if num_val < float(min_pos_esticagem) or num_val > float(max_pos_esticagem):

                        status_pos = "NG"

                        status_ng = True

                        entry.configure(bg="red")  # causador vermelho

                    else:

                        entry.configure(bg="white")

                        if not status_ng:

                            status_pos = "OK"

                except ValueError:

                    status_pos = "NG"

                    status_ng = True

                    entry.configure(bg="red")


        entry_status_pos.config(state="normal")

        entry_status_pos.delete(0, tk.END)

        entry_status_pos.insert(0, status_pos)


        if status_pos == "NG":

            entry_status_pos.config(state="readonly", readonlybackground="red", fg="white")

        elif status_pos == "OK":

            entry_status_pos.config(state="normal", bg="white", fg="black")

        else:

            entry_status_pos.config(state="normal", bg="white", fg="black")


def preencher_parametros_esticagem():

    try:

        cursor = conn.cursor()

        cursor.execute("""

            SELECT TOP 1 minimo, maximo, data, hora

            FROM parametro_esticagem

            ORDER BY id DESC

        """)

        row = cursor.fetchone()

        cursor.close()


        if row:

            minimo, maximo, data_param, hora_param = row

            media = round((minimo + maximo) / 2, 3)

            variacao = round((maximo - minimo) / 2, 3)


            entry_tensao.configure(state="normal")

            entry_tensao.delete(0, tk.END)

            entry_tensao.insert(0, f"{media:.3f}")

            entry_tensao.configure(state="readonly")


            entry_tensao_plusminus.configure(state="normal")

            entry_tensao_plusminus.delete(0, tk.END)

            entry_tensao_plusminus.insert(0, f"{variacao:.3f}")

            entry_tensao_plusminus.configure(state="readonly")


            entry_de.configure(state="normal")

            entry_de.delete(0, tk.END)

            entry_de.insert(0, f"{minimo:.3f}")

            entry_de.configure(state="readonly")


            entry_ate.configure(state="normal")

            entry_ate.delete(0, tk.END)

            entry_ate.insert(0, f"{maximo:.3f}")

            entry_ate.configure(state="readonly")


            # Mostra data + hora vindos do banco

            entry_parametros.configure(state="normal")

            entry_parametros.delete(0, tk.END)

            entry_parametros.insert(

                0,

                f"{data_param.strftime('%d/%m/%Y')} {hora_param.strftime('%H:%M:%S')}"

            )

            entry_parametros.configure(state="readonly")


        else:

            messagebox.showwarning("Aviso", "Parâmetros de esticagem não encontrados.")

    except Exception as e:

        messagebox.showerror("Erro", f"Erro ao buscar parâmetros: {e}")


def preencher_parametros_pos_esticagem():

    try:

        cursor = conn.cursor()

        cursor.execute("""

            SELECT TOP 1 minimo, maximo, data, hora

            FROM parametro_pos_esticagem

            ORDER BY id DESC

        """)

        row = cursor.fetchone()

        cursor.close()


        if row:

            minimo, maximo, data_param, hora_param = row

            media = round((minimo + maximo) / 2, 3)

            variacao = round((maximo - minimo) / 2, 3)


            entry_tensao_3d.configure(state="normal")

            entry_tensao_3d.delete(0, tk.END)

            entry_tensao_3d.insert(0, f"{media:.3f}")

            entry_tensao_3d.configure(state="readonly")


            entry_tensao_plusminus_3d.configure(state="normal")

            entry_tensao_plusminus_3d.delete(0, tk.END)

            entry_tensao_plusminus_3d.insert(0, f"{variacao:.3f}")

            entry_tensao_plusminus_3d.configure(state="readonly")


            entry_de_3d.configure(state="normal")

            entry_de_3d.delete(0, tk.END)

            entry_de_3d.insert(0, f"{minimo:.3f}")

            entry_de_3d.configure(state="readonly")


            entry_ate_3d.configure(state="normal")

            entry_ate_3d.delete(0, tk.END)

            entry_ate_3d.insert(0, f"{maximo:.3f}")

            entry_ate_3d.configure(state="readonly")


            entry_parametros_3d.configure(state="normal")

            entry_parametros_3d.delete(0, tk.END)

            entry_parametros_3d.insert(

                0,

                f"{data_param.strftime('%d/%m/%Y')} {hora_param.strftime('%H:%M:%S')}"

            )

            entry_parametros_3d.configure(state="readonly")

        else:

            messagebox.showwarning("Aviso", "Parâmetros pós-esticagem não encontrados.")

    except Exception as e:

        messagebox.showerror("Erro", f"Erro ao buscar parâmetros pós-esticagem: {e}")


def buscar_opcoes(tabela, coluna):

    if conn is None:

        return []

    try:

        cursor = conn.cursor()

        cursor.execute(f"SELECT {coluna} FROM {tabela}")

        resultados = [row[0] for row in cursor.fetchall()]

        cursor.close()

        return resultados

    except Exception as e:

        print(f"Erro ao buscar dados de {tabela}: {e}")

        return []


def buscar_funcionario_por_matricula(matricula):

    try:

        cursor = conn.cursor()

        cursor.execute("SELECT nome FROM funcionario_nome WHERE reg_func = ?", (matricula,))

        row = cursor.fetchone()

        cursor.close()

        return row[0] if row else None

    except Exception as e:

        print(f"Erro ao buscar funcionário: {e}")

        return None


def extrair_matricula(qr_code):

    """Extrai a matrícula se o QR code terminar com '010101'."""

    import re

    if qr_code.endswith("010101"):

        match = re.match(r'^0*([1-9][0-9]*)', qr_code)

        if match:

            return match.group(1)

    return None  # Retorna None se não encontrar ou não terminar com 010101


def on_matricula_focusout(event):

    widget = event.widget


    # Verifica qual campo disparou o evento

    if widget == entry_func_codigo:

        entry_codigo = entry_func_codigo

        entry_nome = entry_func_nome

        entry_data_target = entry_data

    elif widget == entry_func_codigo_pos:

        entry_codigo = entry_func_codigo_pos

        entry_nome = entry_func_nome_pos

        entry_data_target = entry_data_pos

    else:

        return


    codigo_lido = entry_codigo.get().strip()

    if not codigo_lido:

        entry_nome.configure(state="normal")

        entry_nome.delete(0, tk.END)

        entry_nome.configure(state="readonly")

        entry_data_target.delete(0, tk.END)

        return


    matricula = extrair_matricula(codigo_lido)


    if not matricula:

        messagebox.showwarning("Aviso", f"QR inválido: '{codigo_lido}'")

        entry_codigo.delete(0, tk.END)

        entry_nome.configure(state="normal")

        entry_nome.delete(0, tk.END)

        entry_nome.configure(state="readonly")

        entry_data_target.delete(0, tk.END)

        entry_codigo.focus_set()  # Retorna o foco ao campo para nova digitação

        return


    # Limpa e substitui o campo com a matrícula limpa

    entry_codigo.delete(0, tk.END)

    entry_codigo.insert(0, matricula)


    nome = buscar_funcionario_por_matricula(matricula)

    if nome:

        entry_nome.configure(state="normal")

        entry_nome.delete(0, tk.END)

        entry_nome.insert(0, nome)

        entry_nome.configure(state="readonly")


        hoje = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        entry_data_target.configure(state="normal")

        entry_data_target.delete(0, tk.END)

        entry_data_target.insert(0, hoje)

        entry_data_target.configure(state="readonly")

    else:

        messagebox.showwarning("Aviso", f"Matrícula {matricula} não encontrada.")

        entry_codigo.delete(0, tk.END)

        entry_nome.configure(state="normal")

        entry_nome.delete(0, tk.END)

        entry_nome.configure(state="readonly")

        entry_data_target.delete(0, tk.END)

        entry_codigo.focus_set()


root = tk.Tk()


frame_botoes = tk.Frame(root)

frame_botoes.pack(side="top", anchor="ne", padx=5, pady=5)


# Botão Minimizar

btn_minimizar = tk.Button(frame_botoes, text="_", width=3, command=root.iconify)

btn_minimizar.pack(side="left")


# Botão Fechar

btn_fechar = tk.Button(frame_botoes, text="X", width=3, fg="red", command=root.destroy)

btn_fechar.pack(side="left")

root.title("Esticagem")

root.attributes('-fullscreen', True)

if not db_online:

    def _avisar_modo_offline():

        detalhe = f"\n\nDetalhe: {db_error}" if db_error else ""

        messagebox.showwarning(

            "Modo Offline",

            "Não foi possível conectar ao banco de dados.\nA interface foi aberta em modo offline." + detalhe

        )

    root.after(200, _avisar_modo_offline)


main_frame = tk.Frame(root)

main_frame.pack(fill="both", expand=True)


canvas = tk.Canvas(main_frame)

canvas.pack(side="left", fill="both", expand=True)


scrollbar = tk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)

scrollbar.pack(side="right", fill="y")


canvas.configure(yscrollcommand=scrollbar.set)


form_frame = tk.Frame(canvas)

window = canvas.create_window((0, 0), window=form_frame, anchor="nw")


def on_frame_configure(event=None):

    canvas.configure(scrollregion=canvas.bbox("all"))

    canvas_width = canvas.winfo_width()

    form_width = form_frame.winfo_width()

    x = max((canvas_width - form_width) // 2, 0)

    canvas.coords(window, x, 0)


form_frame.bind("<Configure>", on_frame_configure)

canvas.bind("<Configure>", on_frame_configure)


def _on_mousewheel(event):

    canvas.yview_scroll(int(-1*(event.delta/120)), "units")


canvas.bind_all("<MouseWheel>", _on_mousewheel)


content_frame = tk.Frame(form_frame, width=900)

content_frame.pack(padx=20, pady=20)


label_width = 12

entry_width = 20

font_title = ("Arial", 14, "bold")

font_subtitle = ("Arial", 10)

font_frame_title = ("Arial", 11, "underline")


def add_label_entry(parent, text, row, col_label, col_entry, width=entry_width, readonly=False):

    label = tk.Label(parent, text=text, width=label_width, anchor="e", font=font_subtitle)

    label.grid(row=row, column=col_label, sticky="e", padx=5, pady=5)

    entry = tk.Entry(parent, width=width, font=font_subtitle, justify="right")

    if readonly:

        entry.configure(state="readonly")

    entry.grid(row=row, column=col_entry, sticky="w", padx=5, pady=5)

    return label, entry


titulo = tk.Label(content_frame, text="1. Esticagem", font=font_title)

titulo.grid(row=0, column=0, columnspan=7, pady=(0, 25), sticky="nsew")


_, entry_id_quadro = add_label_entry(content_frame, "ID Quadro (QR):", 1, 0, 1)

entry_id_quadro.configure(justify="left")


# --- FUNÇÕES AUXILIARES PARA ORIGEM ---

def mostrar_origem_como_entry(origem_val):

    """Substitui o combobox por um Entry readonly com o valor da origem"""

    global entry_origem_substituto


    # Se já existir um entry substituto, destrói antes

    if 'entry_origem_substituto' in globals() and entry_origem_substituto.winfo_exists():

        entry_origem_substituto.destroy()


    entry_origem_substituto = tk.Entry(combo_origem.master, font=combo_origem["font"])

    entry_origem_substituto.insert(0, origem_val)

    entry_origem_substituto.config(state="readonly")

    entry_origem_substituto.grid(**combo_origem.grid_position)

    combo_origem.grid_forget()


def mostrar_origem_como_combobox():

    """Restaura o combobox (se o entry tiver sido criado antes)"""

    if 'entry_origem_substituto' in globals() and entry_origem_substituto.winfo_exists():

        entry_origem_substituto.destroy()

    combo_origem.set("")

    combo_origem.grid(**combo_origem.grid_position)


# --- FUNÇÃO PRINCIPAL ---

def validar_e_carregar_quadro(event=None):

    valor = entry_id_quadro.get().strip()


    # --- Validação do QR code ---

    if not valor.endswith("CV050201") or len(valor) <= len("CV050201"):

        messagebox.showerror("Erro", "Código Inválido.")

        entry_id_quadro.delete(0, tk.END)

        entry_id_quadro.focus_set()

        return False


    try:

        cursor = conn.cursor()


        # --- Último NG da pós-esticagem ---

        cursor.execute("""

            SELECT TOP 1 data

            FROM dbo.pos_esticagem

            WHERE id_quadro = ? AND status='NG'

            ORDER BY data DESC

        """, (valor,))

        row_ng = cursor.fetchone()

        data_pos_ng = row_ng[0] if row_ng else None


        # --- Dados da esticagem ---

        cursor.execute("""

            SELECT TOP 1

                id_quadro, origem, disposicao,

                consumo,

                medicao01_x, medicao01_y,

                medicao02_x, medicao02_y,

                medicao03_x, medicao03_y,

                medicao04_x, medicao04_y,

                medicao05_x, medicao05_y,

                data, reg_func, qr_mesh, mesh,

                qr_cola, cola, angulo, observacao, status

            FROM dbo.esticagem

            WHERE id_quadro = ?

            ORDER BY data DESC

        """, (valor,))

        row = cursor.fetchone()


        # --- Limpar campos ---

        entry_id_quadro.config(state="normal")

        entry_id_quadro.delete(0, tk.END)

        entry_id_quadro.insert(0, valor)


        combo_origem.set("")

        combo_origem.config(state="normal")

        disp_var.set("")

        radio_normal.config(state="normal")

        radio_bias.config(state="normal")


        for entry_x, entry_y in entries_tensao:

            entry_x.config(state="normal")

            entry_y.config(state="normal")

            entry_x.delete(0, tk.END)

            entry_y.delete(0, tk.END)


        entry_func_codigo.config(state="normal")

        entry_func_codigo.delete(0, tk.END)

        entry_func_nome.config(state="normal")

        entry_func_nome.delete(0, tk.END)

        entry_func_nome.config(state="readonly")


        entry_bias.config(state="disabled")

        entry_bias.delete(0, tk.END)


        for entry_widget in [entry_mesh_leitura, entry_mesh_info, entry_cola_1, entry_cola_2]:

            entry_widget.config(state="normal")

            entry_widget.delete(0, tk.END)


        entry_motivo.config(state="normal")

        entry_motivo.delete("1.0", tk.END)

        entry_status.config(state="normal")

        entry_status.delete(0, tk.END)

        entry_data.config(state="normal")

        entry_data.delete(0, tk.END)


        # --- Sempre carregar a data da esticagem (readonly) ---

        data_esticagem = row[14] if row else None

        if data_esticagem:

            if isinstance(data_esticagem, datetime):

                entry_data.insert(0, data_esticagem.strftime("%d/%m/%Y %H:%M:%S"))

            else:

                entry_data.insert(0, str(data_esticagem))

        entry_data.config(state="readonly")


        # --- Preencher radio button e Bias ---

        if row:

            disp_val = (row[2] or "").strip().capitalize()

            if disp_val == "Normal":

                disp_var.set("Normal")

                entry_bias.config(state="disabled")

                entry_bias.delete(0, tk.END)

            elif disp_val == "Bias":

                disp_var.set("Bias")

                angulo_val = row[20]

                entry_bias.config(state="normal")

                entry_bias.delete(0, tk.END)

                entry_bias.insert(0, str(angulo_val) if angulo_val is not None else "22.5")

            else:

                disp_var.set("")

                entry_bias.config(state="disabled")

                entry_bias.delete(0, tk.END)

        else:

            disp_var.set("")

            entry_bias.config(state="disabled")

            entry_bias.delete(0, tk.END)


        # --- Preencher campos apenas se esticagem não for NG ou foi refeita após NG ---

        if row:

            origem_val = row[1] or ""

            status_val = row[22]


            if status_val != "NG" and (data_pos_ng is None or data_esticagem > data_pos_ng):

                # --- Bloquear radios e entry.bias somente se houver dados válidos ---

                radio_normal.config(state="disabled")

                radio_bias.config(state="disabled")

                entry_bias.config(state="disabled")


                if origem_val:

                    mostrar_origem_como_entry(origem_val)

                else:

                    mostrar_origem_como_combobox()


                medicoes = row[4:14]

                for i, (entry_x, entry_y) in enumerate(entries_tensao):

                    if medicoes[i*2] is not None:

                        entry_x.insert(0, str(medicoes[i*2]))

                    if medicoes[i*2+1] is not None:

                        entry_y.insert(0, str(medicoes[i*2+1]))

                    entry_x.config(state="readonly")

                    entry_y.config(state="readonly")


                reg_func_val = row[15]

                if reg_func_val:

                    entry_func_codigo.insert(0, str(reg_func_val))

                    nome = buscar_funcionario_por_matricula(str(reg_func_val))

                    entry_func_nome.config(state="normal")

                    entry_func_nome.insert(0, nome)

                    entry_func_nome.config(state="readonly")

                    entry_func_codigo.config(state="disabled")


                if row[16] is not None:

                    entry_mesh_leitura.insert(0, str(row[16]))

                entry_mesh_leitura.config(state="readonly")

                if row[17] is not None:

                    entry_mesh_info.insert(0, str(row[17]))

                entry_mesh_info.config(state="readonly")


                entry_cola_1.insert(0, str(row[18] or ""))

                entry_cola_1.config(state="readonly")

                entry_cola_2.insert(0, str(row[19] or ""))

                entry_cola_2.config(state="readonly")


                if row[21]:

                    entry_motivo.insert("1.0", row[21])

                entry_motivo.config(state="disabled")


                # --- Inserir status e ajustar cor e letra ---

                entry_status.delete(0, tk.END)

                entry_status.insert(0, str(status_val))

                entry_status.config(state="readonly")

                if status_val == "NG":

                    entry_status.config(readonlybackground="red", foreground="white")

                else:

                    entry_status.config(readonlybackground="white", foreground="black")

            else:

                mostrar_origem_como_combobox()

        else:

            mostrar_origem_como_combobox()


        # --- Pós-esticagem ---

        cursor.execute("""

            SELECT TOP 1

                reg_func, espessura_pos_esticagem,

                medicao_01_x, medicao_01_y,

                medicao_02_x, medicao_02_y,

                medicao_03_x, medicao_03_y,

                medicao_04_x, medicao_04_y,

                medicao_05_x, medicao_05_y,

                status, observacao, data

            FROM dbo.pos_esticagem

            WHERE id_quadro = ?

            ORDER BY data DESC

        """, (valor,))

        row_pos = cursor.fetchone()


        if not row_pos or row_pos[12] == "NG":

            limpar_campos_pos_esticagem_editavel()

            entry_media_3d.config(state="normal")

            entry_media_3d.delete(0, tk.END)

        else:

            status_val = row_pos[12]

            reg_func_val, espessura_val, medicoes, observacao_val, data_val = (

                row_pos[0], row_pos[1], row_pos[2:12], row_pos[13], row_pos[14]

            )


            for entry in espessura_3d_entries:

                entry.config(state="readonly")


            entry_func_codigo_pos.config(state="normal")

            entry_func_codigo_pos.delete(0, tk.END)

            if reg_func_val:

                entry_func_codigo_pos.insert(0, str(reg_func_val))

                nome = buscar_funcionario_por_matricula(str(reg_func_val))

                entry_func_nome_pos.config(state="normal")

                entry_func_nome_pos.delete(0, tk.END)

                entry_func_nome_pos.insert(0, nome)

                entry_func_nome_pos.config(state="readonly")

            entry_func_codigo_pos.config(state="disabled")


            for i, (entry_x, entry_y) in enumerate(entries_tensao_3d):

                entry_x.config(state="normal")

                entry_y.config(state="normal")

                x_val = medicoes[i*2]

                y_val = medicoes[i*2+1]

                entry_x.delete(0, tk.END)

                entry_y.delete(0, tk.END)

                if x_val is not None:

                    entry_x.insert(0, str(x_val))

                if y_val is not None:

                    entry_y.insert(0, str(y_val))

                entry_x.config(state="readonly")

                entry_y.config(state="readonly")


            entry_media_3d.config(state="normal")

            entry_media_3d.delete(0, tk.END)

            if espessura_val is not None:

                entry_media_3d.insert(0, str(espessura_val))

            entry_media_3d.config(state="readonly")


            entry_motivo_pos.config(state="normal")

            entry_motivo_pos.delete("1.0", tk.END)

            if observacao_val:

                entry_motivo_pos.insert("1.0", observacao_val)

            entry_motivo_pos.config(state="disabled")


            # --- Status pós-esticagem com cor e letra ---

            entry_status_pos.delete(0, tk.END)

            entry_status_pos.insert(0, str(status_val))

            entry_status_pos.config(state="readonly")

            if status_val == "NG":

                entry_status_pos.config(readonlybackground="red", foreground="white")

            else:

                entry_status_pos.config(readonlybackground="white", foreground="black")


            entry_data_pos.config(state="normal")

            entry_data_pos.delete(0, tk.END)

            if isinstance(data_val, datetime):

                entry_data_pos.insert(0, data_val.strftime("%d/%m/%Y %H:%M:%S"))

            else:

                entry_data_pos.insert(0, str(data_val))

            entry_data_pos.config(state="readonly")


    except Exception as e:

        messagebox.showerror("Erro", f"Falha ao carregar dados do quadro:\n{e}")

        entry_id_quadro.focus_set()

        return False


    return True


def carregar_emulsao_por_quadro(id_quadro):

    """

    Carrega todos os campos de emulsão com base no id_quadro (string).

    Caso o último registro de revelacao_final seja 'NG' + 'Sim',

    só permite carregar emulsão se existir registro de emulsão com data mais recente.

    Caso o último registro de emulsão tenha status 'NG', também não carrega.

    Se id_quadro estiver vazio ou não existir, limpa os campos.

    """

    if not id_quadro or not id_quadro.strip():

        limpar_campos_emulsao_editavel()

        return


    id_quadro = id_quadro.strip()


    try:

        # --- 1. Buscar último registro em revelacao_final ---

        cursor = conn.cursor()

        cursor.execute("""

            SELECT TOP 1 status, reutilizar, data

            FROM dbo.revelacao_final

            WHERE id_quadro = ?

            ORDER BY data DESC

        """, (id_quadro,))

        row_rev = cursor.fetchone()

        cursor.close()


        status_rev = reutilizar_rev = data_rev = None

        if row_rev:

            status_rev, reutilizar_rev, data_rev = row_rev


        # --- 2. Buscar último registro de emulsão ---

        cursor = conn.cursor()

        cursor.execute("""

            SELECT TOP 1

                qr_emulsao, emulsao, polimero,

                espessura_pos_emulsao, espessura_emulsao,

                data, reg_func, status, observacao

            FROM dbo.emulsao

            WHERE id_quadro = ?

            ORDER BY data DESC

        """, (id_quadro,))

        row = cursor.fetchone()

        cursor.close()

        entry_emulsao_qr.focus_set()


        if not row:

            limpar_campos_emulsao_editavel()

            print("DEBUG: Nenhum registro de emulsão encontrado para esse quadro. Campos limpos.")

            return


        qr_emulsao_val, emulsao_val, polimero_val, \
        espessura_pos_emulsao_val, espessura_emulsao_val, \
        data_val, reg_func_val, status_val, observacao_val = row


        # --- 3. Regras de bloqueio (NG + Sim em revelacao_final) ---

        if row_rev and status_rev and status_rev.upper() == "NG" and reutilizar_rev == "Sim":

            if not data_val or (data_rev and data_val <= data_rev):

                # Emulsão não foi refeita ainda → bloquear carregamento

                limpar_campos_emulsao_editavel()

                print("DEBUG: Revelação final NG + Sim e emulsão não refeita → não carregar.")

                return

            else:

                print("DEBUG: Emulsão refeita após revelação final NG+Sim → carregar normalmente.")


        # --- 4. Se último registro de emulsão é NG, não carregar ---

        if status_val and status_val.upper() == "NG":

            print("DEBUG: Último registro de emulsão é NG, nada será carregado.")

            return


        # --- 5. Preencher campos ---

        entry_emulsao_qr.config(state="normal")

        entry_emulsao_qr.delete(0, tk.END)

        entry_emulsao_qr.insert(0, qr_emulsao_val or "")

        entry_emulsao_qr.config(state="readonly")


        entry_emulsao_qr_readonly.config(state="normal")

        entry_emulsao_qr_readonly.delete(0, tk.END)

        entry_emulsao_qr_readonly.insert(0, emulsao_val or "")

        entry_emulsao_qr_readonly.config(state="readonly")


        var_msfilm.set(bool(polimero_val))


        entry_espessura_2.config(state="normal")

        entry_espessura_2.delete(0, tk.END)

        entry_espessura_2.insert(0, str(espessura_pos_emulsao_val or ""))

        entry_espessura_2.config(state="readonly")


        entry_espessura_3.config(state="normal")

        entry_espessura_3.delete(0, tk.END)

        entry_espessura_3.insert(0, str(espessura_emulsao_val or ""))

        entry_espessura_3.config(state="readonly")


        # Código do funcionário

        entry_func_codigo_emulsao.config(state="normal")

        entry_func_codigo_emulsao.delete(0, tk.END)

        if reg_func_val:

            entry_func_codigo_emulsao.insert(0, str(reg_func_val))

        entry_func_codigo_emulsao.config(state="disabled")


        # Nome do funcionário

        entry_func_nome_emulsao.config(state="normal")

        entry_func_nome_emulsao.delete(0, tk.END)

        if reg_func_val:

            nome_func = buscar_funcionario_por_matricula(str(reg_func_val))

            entry_func_nome_emulsao.insert(0, nome_func or "")

        entry_func_nome_emulsao.config(state="readonly")


        entry_status_emulsao.config(state="normal")

        entry_status_emulsao.delete(0, tk.END)

        entry_status_emulsao.insert(0, status_val or "")

        entry_status_emulsao.config(state="readonly")


        entry_motivo_emulsao.config(state="normal")

        entry_motivo_emulsao.delete("1.0", tk.END)

        entry_motivo_emulsao.insert("1.0", observacao_val or "")

        entry_motivo_emulsao.config(state="disabled")


        entry_data_emulsao.config(state="normal")

        entry_data_emulsao.delete(0, tk.END)

        if isinstance(data_val, datetime):

            entry_data_emulsao.insert(0, data_val.strftime("%d/%m/%Y %H:%M:%S"))

        else:

            entry_data_emulsao.insert(0, str(data_val or ""))

        entry_data_emulsao.config(state="readonly")


        print("DEBUG: Dados da emulsão carregados com sucesso.")


    except Exception as e:

        messagebox.showerror("Erro", f"Falha ao carregar dados do quadro ou da emulsão:\n{e}")


label_origem = tk.Label(content_frame, text="Origem:", width=label_width, anchor="e", font=font_subtitle)

label_origem.grid(row=1, column=2, sticky="e", padx=5, pady=5)

origens = buscar_opcoes("origem_BD", "origem")

combo_origem = ttk.Combobox(content_frame, values=origens, width=entry_width-4, font=font_subtitle, state="readonly")

combo_origem.grid(row=1, column=3, sticky="w", padx=5, pady=5)


label_mesh = tk.Label(content_frame, text="Mesh (QR):", width=label_width, anchor="e", font=font_subtitle)

label_mesh.grid(row=3, column=0, sticky="e", padx=5, pady=5)

mesh_frame = tk.Frame(content_frame)

mesh_frame.grid(row=3, column=1, sticky="w", pady=5)

entry_mesh_leitura = tk.Entry(mesh_frame, width=12, font=font_subtitle)

entry_mesh_leitura.grid(row=0, column=0, padx=(0,10))

entry_mesh_info = tk.Entry(mesh_frame, width=25, font=font_subtitle, state="readonly")

entry_mesh_info.grid(row=0, column=1)


def verificar_mesh_info():

    """

    Se o valor no entry_mesh_info for 115,

    força a seleção Normal e desabilita Bias.

    Caso contrário, mantém ambos habilitados.

    """

    valor = entry_mesh_info.get().strip()


    if "115" in valor:  # detecta "115" em qualquer posição

        disp_var.set("Normal")       # força Normal

        radio_bias.config(state="disabled")

        radio_normal.config(state="normal")

    else:

        radio_bias.config(state="normal")

        radio_normal.config(state="normal")


# --- Lista de Mesh permitidos para Importado ---

mesh_importado = [

    "MESH 180 - 155cm",

    "MESH 200 - 155cm",

    "MESH 200 - 115cm",

]


# --- Função para buscar Mesh ---

def buscar_mesh_por_qr(qr):

    if len(qr) < 6:

        return ""


    referencia = qr[:6]

    try:

        referencia_num = int(referencia.lstrip('0'))

    except ValueError:

        return ""


    if not qr[-3:] == "MPI":

        return ""


    try:

        cursor = conn.cursor()

        cursor.execute(

            "SELECT Desc_Proc_01, Tipo FROM mp_sistema WHERE codigo_01 = ?",

            (referencia_num,)

        )

        row = cursor.fetchone()

        cursor.close()


        if row and row[1] == "Mesh":

            return row[0]

        else:

            return ""

    except Exception as e:

        print(f"Erro ao buscar mesh: {e}")

        return ""


# --- Validação do Mesh dependendo da origem ---

def validar_mesh_importado(mesh_valor):

    origem = combo_origem.get()

    if origem == "Importado":

        if mesh_valor not in mesh_importado:

            messagebox.showerror(

                "Erro de validação",

                f"O Mesh '{mesh_valor}' não é permitido para origem Importado.\n\n"

                f"Permitidos: {', '.join(mesh_importado)}"

            )

            # Limpa os campos de Mesh

            entry_mesh_info.configure(state="normal")

            entry_mesh_info.delete(0, tk.END)

            entry_mesh_info.configure(state="readonly")

            entry_mesh_leitura.delete(0, tk.END)

            # Força o foco de volta

            entry_mesh_leitura.after(1, lambda: entry_mesh_leitura.focus_set())

            return False

    return True


# --- Função chamada ao sair do Entry ---

def on_mesh_focusout(event):

    qr = entry_mesh_leitura.get().strip()

    entry_mesh_info.configure(state="normal")

    entry_mesh_info.delete(0, tk.END)


    if qr == "":

        entry_mesh_info.configure(state="readonly")

        return


    mesh_valor = buscar_mesh_por_qr(qr)

    entry_mesh_info.insert(0, mesh_valor if mesh_valor else "Não encontrado")

    entry_mesh_info.configure(state="readonly")


    # Validação da origem

    if mesh_valor:

        validar_mesh_importado(mesh_valor)


        # --- Nova regra: se o mesh tiver "115", só Normal ---

        if "115" in mesh_valor:

            disp_var.set("Normal")

            radio_normal.config(state="normal")

            radio_bias.config(state="disabled")

            # Apaga valor do entry_bias se estava preenchido

            entry_bias.config(state="normal")

            entry_bias.delete(0, tk.END)

            entry_bias.config(state="disabled")

        else:

            radio_normal.config(state="normal")

            radio_bias.config(state="normal")


    # --- Pula automaticamente para entry_cola_1 ---

    entry_cola_1.focus_set()


# --- Função chamada ao mudar a origem ---

def on_origem_selected(event):

    mesh_valor = entry_mesh_info.get().strip()

    if mesh_valor:  # Só valida se já tiver Mesh preenchido

        validar_mesh_importado(mesh_valor)


# Binds

entry_mesh_leitura.bind("<FocusOut>", on_mesh_focusout)


combo_origem.bind("<<ComboboxSelected>>", on_origem_selected)


disp_var = tk.StringVar(value="Normal")

label_disp = tk.Label(content_frame, text="Disposição:", width=label_width, anchor="e", font=font_subtitle)

label_disp.grid(row=3, column=2, sticky="e", padx=5, pady=5)

disp_frame = tk.Frame(content_frame)

disp_frame.grid(row=3, column=3, sticky="w", padx=(5,30), pady=5)


def on_disp_change():

    if disp_var.get() == "Bias":

        entry_bias.configure(state="normal")

        entry_bias.delete(0, tk.END)

        entry_bias.insert(0, "22.5")

        entry_bias.configure(state="readonly")

    else:

        entry_bias.configure(state="normal")

        entry_bias.delete(0, tk.END)

        entry_bias.configure(state="disabled")


radio_normal = tk.Radiobutton(disp_frame, text="Normal", variable=disp_var, value="Normal", command=on_disp_change, font=font_subtitle)

radio_normal.pack(side="left", padx=(0,20))

radio_bias = tk.Radiobutton(disp_frame, text="Bias", variable=disp_var, value="Bias", command=on_disp_change, font=font_subtitle)

radio_bias.pack(side="left", padx=(0,10))

entry_bias = tk.Entry(disp_frame, width=18, font=font_subtitle, state="disabled", justify="right")

entry_bias.pack(side="left")


frame_aplicacao_cola = tk.LabelFrame(content_frame, text="Aplicação da cola", font=font_frame_title, padx=10, pady=10)

frame_aplicacao_cola.grid(row=4, column=0, columnspan=3, sticky="nsew", padx=(10,20), pady=(20,10))


frame_tensao = tk.LabelFrame(content_frame, text="Tabela de Tensão (mm) X e Y", font=font_frame_title, padx=10, pady=10)

frame_tensao.grid(row=4, column=3, columnspan=4, sticky="nsew", padx=(20,10), pady=(20,10))


for col in range(4):

    frame_aplicacao_cola.grid_columnconfigure(col, weight=1)

for col in range(3):

    frame_tensao.grid_columnconfigure(col, weight=1)


tk.Label(frame_aplicacao_cola, text="Tensão (mm):", font=font_subtitle).grid(row=0, column=0, sticky="e", pady=3)

entry_tensao = tk.Entry(frame_aplicacao_cola, width=12, font=font_subtitle, state="readonly", justify="right")

entry_tensao.grid(row=0, column=1, pady=3, sticky="w")

label_pm = tk.Label(frame_aplicacao_cola, text="+/-", font=font_subtitle)

label_pm.grid(row=0, column=2, pady=3, sticky="nsew")

entry_tensao_plusminus = tk.Entry(frame_aplicacao_cola, width=12, font=font_subtitle, state="readonly", justify="right")

entry_tensao_plusminus.grid(row=0, column=3, pady=3, sticky="w")


tk.Label(frame_aplicacao_cola, text="De:", font=font_subtitle).grid(row=1, column=0, sticky="e", pady=3)

entry_de = tk.Entry(frame_aplicacao_cola, width=12, font=font_subtitle, state="readonly", justify="right")

entry_de.grid(row=1, column=1, pady=3, sticky="w")

label_ate = tk.Label(frame_aplicacao_cola, text="à", font=font_subtitle)

label_ate.grid(row=1, column=2, pady=3, sticky="nsew")

entry_ate = tk.Entry(frame_aplicacao_cola, width=12, font=font_subtitle, state="readonly", justify="right")

entry_ate.grid(row=1, column=3, pady=3, sticky="w")


tk.Label(frame_aplicacao_cola, text="Parâmetros de:", font=font_subtitle).grid(row=2, column=0, sticky="e", pady=3)

entry_parametros = tk.Entry(frame_aplicacao_cola, width=25, font=font_subtitle, state="readonly")

entry_parametros.grid(row=2, column=1, pady=3, sticky="w")


tk.Label(frame_aplicacao_cola, text="Cola (QR):", font=font_subtitle).grid(row=3, column=0, sticky="e", pady=3)

entry_cola_1 = tk.Entry(frame_aplicacao_cola, width=12, font=font_subtitle)

entry_cola_1.grid(row=3, column=1, pady=3, sticky="w")

entry_cola_2 = tk.Entry(frame_aplicacao_cola, width=30, font=font_subtitle, state="readonly")

entry_cola_2.grid(row=3, column=2, columnspan=2, pady=3, sticky="w")


def buscar_desc_proc_por_qr(qr):

    """

    Extrai o número de referência, verifica sufixo 'MPI' e se Tipo = 'Cola'.

    Retorna o valor da coluna Desc_Proc_01 da tabela mp_sistema.

    """

    if len(qr) < 6:

        return None


    referencia = qr[:6]

    try:

        referencia_num = int(referencia.lstrip('0'))

    except ValueError:

        return None


    # Verificar se os últimos 3 caracteres são 'MPI'

    if not qr[-3:].upper() == "MPI":

        return None


    # Buscar no banco

    try:

        cursor = conn.cursor()

        cursor.execute(

            "SELECT Desc_Proc_01, Tipo FROM mp_sistema WHERE codigo_01 = ?",

            (referencia_num,)

        )

        row = cursor.fetchone()

        cursor.close()


        if row and row[1] and row[1].strip().lower() == "cola":

            return row[0]  # Retorna Desc_Proc_01

        else:

            return None

    except Exception as e:

        print(f"Erro ao buscar Desc_Proc_01: {e}")

        return None


def on_cola_focusout(event):

    qr = entry_cola_1.get().strip()

    entry_cola_2.configure(state="normal")

    entry_cola_2.delete(0, tk.END)


    if qr == "":

        entry_cola_2.configure(state="readonly")

        return


    valor = buscar_desc_proc_por_qr(qr)

    if valor is not None:

        entry_cola_2.insert(0, valor)

    else:

        entry_cola_2.insert(0, "Não encontrado")


    entry_cola_2.configure(state="readonly")


entry_cola_1.bind("<FocusOut>", on_cola_focusout)


tk.Label(frame_tensao, text="Tensão (mm):", font=font_subtitle).grid(row=0, column=0, padx=5)

tk.Label(frame_tensao, text="X", font=font_subtitle).grid(row=0, column=1, padx=10)

tk.Label(frame_tensao, text="Y", font=font_subtitle).grid(row=0, column=2, padx=10)


entries_tensao = []

for i in range(5):

    tk.Label(frame_tensao, text=str(i+1), font=font_subtitle).grid(row=i+1, column=0, padx=5, pady=2)

    entry_x = tk.Entry(frame_tensao, width=8, justify="right")

    entry_x.grid(row=i+1, column=1, padx=5, pady=2)

    entry_y = tk.Entry(frame_tensao, width=8, justify="right")

    entry_y.grid(row=i+1, column=2, padx=5, pady=2)

    entries_tensao.append((entry_x, entry_y))

for entry_x, entry_y in entries_tensao:

    entry_x.bind("<FocusOut>", validar_entry)

    entry_y.bind("<FocusOut>", validar_entry)


func_data_frame = tk.Frame(content_frame)

func_data_frame.grid(row=5, column=0, columnspan=7, pady=(20,15), sticky="ew")


for col in range(5):

    func_data_frame.grid_columnconfigure(col, weight=1)


label_func_codigo = tk.Label(func_data_frame, text="Funcionário:", width=label_width, anchor="e", font=font_subtitle)

label_func_codigo.grid(row=0, column=0, padx=5, pady=5, sticky="e")

entry_func_codigo = tk.Entry(func_data_frame, width=12, font=font_subtitle)

entry_func_codigo.grid(row=0, column=1, padx=5, pady=5, sticky="w")

entry_func_codigo.bind("<FocusOut>", on_matricula_focusout)


entry_func_nome = tk.Entry(func_data_frame, width=35, font=font_subtitle, state="readonly")

entry_func_nome.grid(row=0, column=2, padx=5, pady=5, sticky="w")


label_data = tk.Label(func_data_frame, text="Data:", width=label_width, anchor="e", font=font_subtitle)

label_data.grid(row=0, column=3, padx=5, pady=5, sticky="e")

entry_data = tk.Entry(func_data_frame, width=25, font=font_subtitle,state="readonly")

entry_data.grid(row=0, column=4, padx=5, pady=5, sticky="w")


frame_apontamento = tk.LabelFrame(content_frame, text="Apontamento - Esticagem", font=font_frame_title, padx=10, pady=10)

frame_apontamento.grid(row=6, column=0, columnspan=7, sticky="ew", padx=10, pady=(0,20))


for col in range(7):

    frame_apontamento.grid_columnconfigure(col, weight=1)


tk.Label(frame_apontamento, text="Status:", font=font_subtitle, width=label_width, anchor="e").grid(row=0, column=0, padx=5, pady=5, sticky="e")

entry_status = tk.Entry(frame_apontamento, width=15, font=font_subtitle, justify="right")

entry_status.grid(row=0, column=1, padx=5, pady=5, sticky="w")


tk.Label(frame_apontamento, text="Observação:", font=font_subtitle, width=label_width, anchor="e").grid(row=1, column=0, padx=5, pady=5, sticky="ne")

entry_motivo = tk.Text(frame_apontamento, width=50, height=4, font=font_subtitle)

entry_motivo.grid(row=1, column=1, columnspan=5, padx=5, pady=5, sticky="w")


botao_salvar_esticagem = tk.Button(

    frame_apontamento,

    text="Salvar Esticagem",

    font=font_subtitle,


)

botao_salvar_esticagem.grid(row=2, column=0, columnspan=7, pady=(10,0))


# --- Função para salvar dados de esticagem ---

def salvar_dados_esticagem():

    try:

        cursor = conn.cursor()


        # --- Extrair valores dos campos ---

        id_quadro = entry_id_quadro.get().strip()

        origem = combo_origem.get().strip()

        disposicao = disp_var.get()

        angulo = 22.5 if disposicao == "Bias" else None

        reg_func = entry_func_codigo.get().strip()

        qr_mesh = entry_mesh_leitura.get().strip()

        mesh = entry_mesh_info.get().strip()

        qr_cola = entry_cola_1.get().strip()

        cola = entry_cola_2.get().strip()

        status = entry_status.get().strip()

        observacao = entry_motivo.get("1.0", tk.END).strip()


        # --- Verificar duplicidade ---

        # Existe registro OK em esticagem?

        cursor.execute("""

            SELECT COUNT(*)

            FROM dbo.esticagem

            WHERE id_quadro = ? AND status = 'OK'

        """, (id_quadro,))

        count_esticagem_ok = cursor.fetchone()[0]


        # Verificar último registro em pos_esticagem

        cursor.execute("""

            SELECT TOP 1 status

            FROM dbo.pos_esticagem

            WHERE id_quadro = ?

            ORDER BY data DESC

        """, (id_quadro,))

        row_pos = cursor.fetchone()

        status_pos = row_pos[0] if row_pos else None


        # --- Decidir se pode salvar ---

        if count_esticagem_ok > 0:

            if status_pos == "NG":

                # Reset liberado → pode salvar novamente

                pass

            else:

                # Bloqueia se já existe OK e não houve reset

                messagebox.showwarning("Aviso", f"Já existe um registro OK para o quadro {id_quadro}.")

                return


        if status_pos == "OK":

            # Se pós-esticagem está OK, também não pode salvar

            messagebox.showwarning("Aviso", f"O quadro {id_quadro} já foi aprovado na Pós-Esticagem.")

            return


        # --- Medições X/Y ---

        medicoes = []

        for entry_x, entry_y in entries_tensao:

            x_val = entry_x.get().strip()

            y_val = entry_y.get().strip()

            medicoes.append((x_val, y_val))


        # --- Validação obrigatória ---

        campos_obrigatorios = {

            "ID Quadro": id_quadro,

            "Origem": origem,

            "Funcionário": reg_func,

            "QR Mesh": qr_mesh,

            "Mesh": mesh,

            "QR Cola": qr_cola,

            "Cola": cola,

            "Status": status

        }


        # Verificar entradas de medição

        for i, (x_val, y_val) in enumerate(medicoes, start=1):

            if not x_val or not y_val:

                messagebox.showwarning("Aviso", f"Preencha todas as medições (Medição {i}).")

                return


        # Checar os campos obrigatórios

        for nome, valor in campos_obrigatorios.items():

            if not valor:

                messagebox.showwarning("Aviso", f"Preencha o campo obrigatório: {nome}.")

                return


        # Observação obrigatória somente se status for NG

        if status.upper() == "NG" and not observacao:

            messagebox.showwarning("Aviso", "Campo Observação é obrigatório quando o status é NG.")

            return


        # --- Preparar valores para INSERT (data com GETDATE()) ---

        sql = """

        INSERT INTO dbo.esticagem (

            id_quadro, origem, disposicao,

            medicao01_x, medicao01_y,

            medicao02_x, medicao02_y,

            medicao03_x, medicao03_y,

            medicao04_x, medicao04_y,

            medicao05_x, medicao05_y,

            data, reg_func, qr_mesh, mesh, qr_cola, cola,

            angulo, observacao, status

        ) VALUES (?, ?, ?,

                  ?, ?,

                  ?, ?,

                  ?, ?,

                  ?, ?,

                  ?, ?,

                  GETDATE(), ?, ?, ?, ?, ?,

                  ?, ?, ?)

        """


        valores_insert = [

            id_quadro, origem, disposicao,

            medicoes[0][0], medicoes[0][1],

            medicoes[1][0], medicoes[1][1],

            medicoes[2][0], medicoes[2][1],

            medicoes[3][0], medicoes[3][1],

            medicoes[4][0], medicoes[4][1],

            reg_func, qr_mesh, mesh, qr_cola, cola,

            angulo, observacao, status

        ]


        cursor.execute(sql, valores_insert)

        conn.commit()

        cursor.close()

        messagebox.showinfo("Sucesso", "Dados salvos com sucesso!")


        # --- Resetar todos os campos e cores ---

        def reset_entry(entry_widget):

            entry_widget.config(state="normal", foreground="black", readonlybackground="white")

            entry_widget.delete(0, tk.END)


        reset_entry(entry_id_quadro)

        combo_origem.set('')


        # Disposição

        disp_var.set("Normal")

        entry_bias.configure(state="disabled")

        entry_bias.delete(0, tk.END)


        reset_entry(entry_func_codigo)

        entry_func_nome.configure(state="normal")

        entry_func_nome.delete(0, tk.END)

        entry_func_nome.configure(state="readonly")


        reset_entry(entry_mesh_leitura)

        entry_mesh_info.configure(state="normal")

        entry_mesh_info.delete(0, tk.END)

        entry_mesh_info.configure(state="readonly")


        reset_entry(entry_cola_1)

        entry_cola_2.configure(state="normal")

        entry_cola_2.delete(0, tk.END)

        entry_cola_2.configure(state="readonly")


        entry_data.configure(state="normal")

        entry_data.delete(0, tk.END)

        entry_data.configure(state="readonly")


        entry_status.configure(state="normal")

        entry_status.delete(0, tk.END)

        entry_status.configure(foreground="black", readonlybackground="white")


        entry_motivo.delete("1.0", tk.END)


        for entry_x, entry_y in entries_tensao:

            reset_entry(entry_x)

            reset_entry(entry_y)


        # Focar novamente no ID Quadro

        entry_id_quadro.focus_set()


    except Exception as e:

        messagebox.showerror("Erro", f"Falha ao salvar dados: {e}")


def limpar_campo_esticagem():

    # Limpar campos imputáveis e dependentes

    entry_id_quadro.delete(0, tk.END)

    combo_origem.set('')


    # Disposição

    disp_var.set("Normal")

    entry_bias.configure(state="disabled")

    entry_bias.delete(0, tk.END)


    entry_func_codigo.delete(0, tk.END)

    entry_func_nome.configure(state="normal")

    entry_func_nome.delete(0, tk.END)

    entry_func_nome.configure(state="readonly")


    entry_mesh_leitura.delete(0, tk.END)

    entry_mesh_info.configure(state="normal")

    entry_mesh_info.delete(0, tk.END)

    entry_mesh_info.configure(state="readonly")


    entry_cola_1.delete(0, tk.END)

    entry_cola_2.configure(state="normal")

    entry_cola_2.delete(0, tk.END)

    entry_cola_2.configure(state="readonly")


    entry_data.configure(state="normal")

    entry_data.delete(0, tk.END)

    entry_data.configure(state="readonly")


    entry_status.delete(0, tk.END)

    entry_motivo.delete("1.0", tk.END)


    for entry_x, entry_y in entries_tensao:

        entry_x.delete(0, tk.END)

        entry_y.delete(0, tk.END)


    # Focar novamente no ID Quadro

    entry_id_quadro.focus_set()


# --- Vincular função ao botão existente ---

botao_salvar_esticagem.configure(command=salvar_dados_esticagem)


titulo_pos_esticagem = tk.Label(content_frame, text="2. Pós-esticagem", font=font_title)

titulo_pos_esticagem.grid(row=7, column=0, columnspan=7, pady=(30, 25), sticky="nsew")


frame_3dias_cola = tk.LabelFrame(content_frame, text="3 dias após a aplicação da cola", font=font_frame_title, padx=10, pady=10)

frame_3dias_cola.grid(row=8, column=0, columnspan=3, sticky="nsew", padx=(10,20), pady=(10,10))


frame_tensao_pos = tk.LabelFrame(content_frame, text="Tabela de Tensão (mm) X e Y", font=font_frame_title, padx=10, pady=10)

frame_tensao_pos.grid(row=8, column=3, columnspan=4, sticky="nsew", padx=(20,10), pady=(10,10))


for col in range(4):

    frame_3dias_cola.grid_columnconfigure(col, weight=1)

for col in range(3):

    frame_tensao_pos.grid_columnconfigure(col, weight=1)


def create_3dias_entry(parent, row, col, width=12, readonly=True):

    e = tk.Entry(parent, width=width, font=font_subtitle, justify="right")

    if readonly:

        e.configure(state="readonly")

    e.grid(row=row, column=col, pady=3, sticky="w")

    return e


# Linha 0 — Tensão (mm)

tk.Label(frame_3dias_cola, text="Tensão (mm):", font=font_subtitle).grid(row=0, column=0, sticky="e", pady=3)


entry_tensao_3d = create_3dias_entry(frame_3dias_cola, 0, 1)


label_pm_3d = tk.Label(frame_3dias_cola, text="+/-", font=font_subtitle)

label_pm_3d.grid(row=0, column=2, pady=3)


entry_tensao_plusminus_3d = create_3dias_entry(frame_3dias_cola, 0, 3)


# Linha 1 — De/Até

tk.Label(frame_3dias_cola, text="De:", font=font_subtitle).grid(row=1, column=0, sticky="e", pady=3)

entry_de_3d = create_3dias_entry(frame_3dias_cola, 1, 1)


label_ate_3d = tk.Label(frame_3dias_cola, text="à", font=font_subtitle)

label_ate_3d.grid(row=1, column=2, pady=3, sticky="nsew")


entry_ate_3d = create_3dias_entry(frame_3dias_cola, 1, 3)


entry_ate_3d = create_3dias_entry(frame_3dias_cola, 1, 3)


tk.Label(frame_3dias_cola, text="Parâmetros de:", font=font_subtitle).grid(

    row=2, column=0, sticky="e", pady=3

)

entry_parametros_3d = create_3dias_entry(frame_3dias_cola, 2, 1)

entry_parametros_3d.config(

    width=25,

    font=font_subtitle,

    state="readonly",

    justify="left"

)


tk.Label(frame_3dias_cola, text="Espessura (μm):", font=font_subtitle).grid(

    row=3, column=0, sticky="e", pady=3

)

frame_espessura_3d = tk.Frame(frame_3dias_cola)

frame_espessura_3d.grid(row=3, column=1, sticky="w", pady=3)


espessura_3d_vars = [tk.StringVar() for _ in range(5)]


# 1. Cria a variável que vai armazenar a média

media_3d_var = tk.StringVar()


# 2. Cria os Entry de input e adiciona trace

espessura_3d_vars = [tk.StringVar() for _ in range(5)]


def calcular_media_espessura_3d(*args):

    try:

        valores = [float(var.get()) for var in espessura_3d_vars if var.get().strip() != ""]

        if valores:

            media = sum(valores) / len(valores)

            media_3d_var.set(str(int(round(media))))

        else:

            media_3d_var.set("")

    except ValueError:

        media_3d_var.set("")


# Entry de input

for i, var in enumerate(espessura_3d_vars):

    entry = tk.Entry(frame_espessura_3d, textvariable=var, width=5)

    entry.grid(row=0, column=i, padx=2)

    var.trace_add("write", calcular_media_espessura_3d)


# Entry readonly da média (agora media_3d_var já existe)

entry_media_3d = tk.Entry(frame_espessura_3d, textvariable=media_3d_var, width=10, state="readonly")

entry_media_3d.grid(row=1, column=0, columnspan=5, pady=5)


tk.Label(frame_tensao_pos, text="Tensão (mm):", font=font_subtitle).grid(row=0, column=0, padx=5)

tk.Label(frame_tensao_pos, text="X", font=font_subtitle).grid(row=0, column=1, padx=10)

tk.Label(frame_tensao_pos, text="Y", font=font_subtitle).grid(row=0, column=2, padx=10)


entries_tensao_3d = []

for i in range(5):

    tk.Label(frame_tensao_pos, text=str(i+1), font=font_subtitle).grid(row=i+1, column=0, padx=5, pady=2)

    entry_x = tk.Entry(frame_tensao_pos, width=8, justify="right")

    entry_x.grid(row=i+1, column=1, padx=5, pady=2)

    entry_y = tk.Entry(frame_tensao_pos, width=8, justify="right")

    entry_y.grid(row=i+1, column=2, padx=5, pady=2)

    entries_tensao_3d.append((entry_x, entry_y))

for entry_x, entry_y in entries_tensao_3d:

    entry_x.bind("<FocusOut>", validar_entry)

    entry_y.bind("<FocusOut>", validar_entry)


espessura_3d_entries = []

for i, var in enumerate(espessura_3d_vars):

    entry = tk.Entry(frame_espessura_3d, textvariable=var, width=5, font=font_subtitle, justify="center")

    entry.grid(row=0, column=i, padx=2)

    var.trace_add("write", calcular_media_espessura_3d)

    espessura_3d_entries.append(entry)


entry_media_espessura = tk.Entry(frame_espessura_3d, textvariable=media_3d_var, width=10, font=font_subtitle, justify="center", state="readonly")

entry_media_espessura.grid(row=1, column=0, columnspan=5, pady=5)


func_data_pos_frame = tk.Frame(content_frame)

func_data_pos_frame.grid(row=9, column=0, columnspan=7, pady=(20,15), sticky="ew")


for col in range(5):

    func_data_pos_frame.grid_columnconfigure(col, weight=1)


label_func_codigo_pos = tk.Label(func_data_pos_frame, text="Funcionário:", width=label_width, anchor="e", font=font_subtitle)

label_func_codigo_pos.grid(row=0, column=0, padx=5, pady=5, sticky="e")

entry_func_codigo_pos = tk.Entry(func_data_pos_frame, width=12, font=font_subtitle)

entry_func_codigo_pos.grid(row=0, column=1, padx=5, pady=5, sticky="w")


entry_func_nome_pos = tk.Entry(func_data_pos_frame, width=35, font=font_subtitle, state="readonly")

entry_func_nome_pos.grid(row=0, column=2, padx=5, pady=5, sticky="w")


label_data_pos = tk.Label(func_data_pos_frame, text="Data:", width=label_width, anchor="e", font=font_subtitle)

label_data_pos.grid(row=0, column=3, padx=5, pady=5, sticky="e")

entry_data_pos = tk.Entry(func_data_pos_frame, width=25, font=font_subtitle, state="readonly")

entry_data_pos.grid(row=0, column=4, padx=5, pady=5, sticky="w",)


def on_matricula_pos_focusout(event):

    matricula = entry_func_codigo_pos.get().strip()

    if matricula == "":

        entry_func_nome_pos.configure(state="normal")

        entry_func_nome_pos.delete(0, tk.END)

        entry_func_nome_pos.configure(state="readonly")

        entry_data_pos.delete(0, tk.END)

        return


    nome = buscar_funcionario_por_matricula(matricula)

    if nome:

        entry_func_nome_pos.configure(state="normal")

        entry_func_nome_pos.delete(0, tk.END)

        entry_func_nome_pos.insert(0, nome)

        entry_func_nome_pos.configure(state="readonly")


        hoje = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        entry_data_pos.configure(state="normal")

        entry_data_pos.delete(0, tk.END)

        entry_data_pos.insert(0, hoje)

        entry_data_pos.configure(state="readonly")

    else:

        messagebox.showwarning("Aviso", f"Matrícula {matricula} não encontrada.")

        entry_func_nome_pos.configure(state="normal")

        entry_func_nome_pos.delete(0, tk.END)

        entry_func_nome_pos.configure(state="readonly")

        entry_data_pos.delete(0, tk.END)


entry_func_codigo.bind("<FocusOut>", on_matricula_focusout)

entry_func_codigo_pos.bind("<FocusOut>", on_matricula_focusout)


frame_apontamento_pos = tk.LabelFrame(content_frame, text="Apontamento - Pós-Esticagem", font=font_frame_title, padx=10, pady=10)

frame_apontamento_pos.grid(row=10, column=0, columnspan=7, sticky="ew", padx=10, pady=(0, 20))


for col in range(7):

    frame_apontamento_pos.grid_columnconfigure(col, weight=1)


tk.Label(frame_apontamento_pos, text="Status:", font=font_subtitle, width=label_width, anchor="e").grid(row=0, column=0, padx=5, pady=5, sticky="e")

entry_status_pos = tk.Entry(frame_apontamento_pos, width=15, font=font_subtitle, justify="right")

entry_status_pos.grid(row=0, column=1, padx=5, pady=5, sticky="w")


tk.Label(frame_apontamento_pos, text="Observação:", font=font_subtitle, width=label_width, anchor="e").grid(row=1, column=0, padx=5, pady=5, sticky="ne")

entry_motivo_pos = tk.Text(frame_apontamento_pos, width=50, height=4, font=font_subtitle)

entry_motivo_pos.grid(row=1, column=1, columnspan=5, padx=5, pady=5, sticky="w")


def salvar_e_limpar_pos_esticagem():

    try:

        cursor = conn.cursor()


        id_quadro_val = entry_id_quadro.get().strip()

        if not id_quadro_val:

            messagebox.showwarning("Aviso", "Informe o ID do Quadro antes de salvar!")

            return


        # --- Verificar se o quadro existe em esticagem ---

        cursor.execute("SELECT COUNT(*) FROM dbo.esticagem WHERE id_quadro = ?", (id_quadro_val,))

        existe = cursor.fetchone()[0]


        if existe == 0:

            messagebox.showerror("Erro", "É necessário preencher Esticagem primeiro.")


            # --- Limpar também os campos de pós-esticagem ---

            for entry_x, entry_y in entries_tensao_3d:

                entry_x.delete(0, tk.END)

                entry_y.delete(0, tk.END)

            for entry in espessura_3d_entries:

                entry.delete(0, tk.END)

            entry_media_3d.delete(0, tk.END)

            entry_status_pos.delete(0, tk.END)

            entry_motivo_pos.delete("1.0", tk.END)

            entry_func_codigo_pos.delete(0, tk.END)

            entry_func_nome_pos.delete(0, tk.END)

            entry_data_pos.delete(0, tk.END)


            return


        # --- Preparar dados ---

        reg_func_val = entry_func_codigo_pos.get().strip()

        status_val = entry_status_pos.get().strip()

        observacao_val = entry_motivo_pos.get("1.0", tk.END).strip()


        # --- Validar obrigatoriedade ---

        if not reg_func_val or not status_val:

            messagebox.showwarning("Aviso", "Preencha todos os campos obrigatórios!")

            return


        # --- Verificar duplicidade em pos_esticagem ---

        cursor.execute("""

            SELECT TOP 1 status

            FROM dbo.pos_esticagem

            WHERE id_quadro = ?

            ORDER BY data DESC

        """, (id_quadro_val,))

        row_pos = cursor.fetchone()

        ultimo_status_pos = row_pos[0] if row_pos else None


        if ultimo_status_pos == "OK" and status_val.upper() == "OK":

            messagebox.showwarning("Aviso", f"Já existe um registro OK para o quadro {id_quadro_val}. Não é permitido sobrescrever.")

            return


        # Validar todas as medições

        medicoes_val = []

        for i, (entry_x, entry_y) in enumerate(entries_tensao_3d, start=1):

            x = entry_x.get().strip()

            y = entry_y.get().strip()

            if not x or not y:

                messagebox.showwarning("Aviso", f"Preencha todas as medições (linha {i}).")

                return

            try:

                medicoes_val.extend([float(x), float(y)])

            except ValueError:

                messagebox.showwarning("Aviso", f"Valores inválidos na medição {i}.")

                return


        media_val = media_3d_var.get().strip()

        if not media_val:

            messagebox.showwarning("Aviso", "Informe a média da espessura.")

            return

        try:

            espessura_val = int(media_val)

        except ValueError:

            messagebox.showwarning("Aviso", "Média de espessura inválida.")

            return


        # Observação obrigatória apenas quando status = "NG"

        if status_val.upper() == "NG" and not observacao_val:

            messagebox.showwarning("Aviso", "Observação é obrigatória quando o status for NG.")

            return


        data_val = datetime.now()


        # --- Inserir pós-esticagem ---

        cursor.execute("""

            INSERT INTO dbo.pos_esticagem (

                id_quadro, reg_func, espessura_pos_esticagem,

                medicao_01_x, medicao_01_y,

                medicao_02_x, medicao_02_y,

                medicao_03_x, medicao_03_y,

                medicao_04_x, medicao_04_y,

                medicao_05_x, medicao_05_y,

                status, observacao, data

            )

            VALUES (?, ?, ?,

                    ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)

        """, (

            id_quadro_val,

            int(reg_func_val),

            espessura_val,

            *medicoes_val,

            status_val,

            observacao_val if observacao_val else None,

            data_val

        ))

        conn.commit()


        # --- Limpar todos os campos ---

        for widget in [

            entry_id_quadro, entry_bias, entry_func_codigo, entry_func_nome,

            entry_mesh_leitura, entry_mesh_info, entry_cola_1, entry_cola_2,

            entry_data, entry_status, entry_func_codigo_pos, entry_func_nome_pos,

            entry_data_pos, entry_media_3d, entry_status_pos

        ]:

            widget.config(state="normal")

            widget.delete(0, tk.END)


        combo_origem.set('')

        disp_var.set("Normal")

        entry_motivo.delete("1.0", tk.END)

        entry_motivo_pos.delete("1.0", tk.END)


        for entry_x, entry_y in entries_tensao:

            entry_x.config(state="normal")

            entry_y.config(state="normal")

            entry_x.delete(0, tk.END)

            entry_y.delete(0, tk.END)


        for entry_x, entry_y in entries_tensao_3d:

            entry_x.config(state="normal")

            entry_y.config(state="normal")

            entry_x.delete(0, tk.END)

            entry_y.delete(0, tk.END)


        for entry in espessura_3d_entries:

            entry.config(state="normal")

            entry.delete(0, tk.END)


        radio_normal.config(state="normal")

        radio_bias.config(state="normal")


        messagebox.showinfo("Sucesso", "Dados de pós-esticagem salvos com sucesso!")


        # Voltar foco no ID quadro

        entry_id_quadro.focus_set()


    except Exception as e:

        messagebox.showerror("Erro", f"Falha ao salvar e limpar campos: {e}")


botao_salvar_pos = tk.Button(

    frame_apontamento_pos,

    text="Salvar Pós-Esticagem",

    font=font_subtitle,

    width=20,

    command=salvar_e_limpar_pos_esticagem

)

botao_salvar_pos.grid(row=2, column=0, columnspan=7, pady=(10, 0))


titulo_emulsao = tk.Label(content_frame, text="3. Emulsão", font=font_title)

titulo_emulsao.grid(row=11, column=0, columnspan=7, pady=(30, 25), sticky="nsew")


font_subtitle = ("Arial", 10)

label_width = 15


titulo_emulsao = tk.Label(content_frame, text="3. Emulsão", font=font_title)

titulo_emulsao.grid(row=11, column=0, columnspan=7, pady=(30, 25), sticky="nsew")


frame_emulsao_interno = tk.Frame(content_frame)

frame_emulsao_interno.grid(row=12, column=0, columnspan=7, sticky="ew", padx=0, pady=5)


for col in range(7):

    frame_emulsao_interno.columnconfigure(col, weight=0)


label_emulsao_qr = tk.Label(

    frame_emulsao_interno,

    text="Emulsão (QR):",

    font=font_subtitle,

    width=12,

    anchor="e"

)

label_emulsao_qr.grid(row=0, column=0, padx=(0,5), pady=5, sticky="e")


entry_emulsao_qr = tk.Entry(

    frame_emulsao_interno,

    width=18,

    font=font_subtitle,

    justify="left"

)

entry_emulsao_qr.grid(row=0, column=1, padx=(0,15), pady=5, sticky="w")


entry_emulsao_qr_readonly = tk.Entry(

    frame_emulsao_interno,

    width=18,

    font=font_subtitle,

    justify="left",

    state="readonly"

)

entry_emulsao_qr_readonly.grid(row=0, column=2, padx=(0,15), pady=5, sticky="w")


label_polimero = tk.Label(

    frame_emulsao_interno,

    text="Polímero:",

    font=font_subtitle,

    width=10,

    anchor="e"

)

label_polimero.grid(row=0, column=3, padx=(0,5), pady=5, sticky="e")


var_msfilm = tk.BooleanVar()


checkbox_msfilm = tk.Checkbutton(

    frame_emulsao_interno,

    text="MsFilm",

    variable=var_msfilm,

    font=font_subtitle

)

checkbox_msfilm.grid(row=0, column=4, padx=(0,10), pady=5, sticky="w")


checkbox_msfilm.bind("<Button-1>", lambda e: "break")

checkbox_msfilm.bind("<space>", lambda e: "break")

checkbox_msfilm.bind("<Return>", lambda e: "break")


def verificar_msfilm(*args):

    valor = entry_emulsao_qr_readonly.get().strip()

    var_msfilm.set("BC-10" in valor)


entry_emulsao_qr_readonly_var = tk.StringVar()

entry_emulsao_qr_readonly.configure(textvariable=entry_emulsao_qr_readonly_var)

entry_emulsao_qr_readonly_var.trace_add("write", verificar_msfilm)


for col in range(7):

    frame_emulsao_interno.columnconfigure(col, weight=1)


import tkinter.font as font

font_subtitle_underline = font.Font(family="Arial", size=11, underline=True)


frame_parametro_espessura = tk.LabelFrame(

    frame_emulsao_interno,

    text="Parâmetro Espessura",

    font=font_subtitle_underline,

    padx=10,

    pady=10

)

frame_parametro_espessura.grid(row=1, column=0, columnspan=7, sticky="ew", padx=0, pady=5)


for col in range(7):

    frame_parametro_espessura.columnconfigure(col, weight=1)


tk.Label(frame_parametro_espessura, text="Espessura (µm):", font=font_subtitle)\
    .grid(row=0, column=1, sticky="e", pady=3)


entry_espessura_base = tk.Entry(frame_parametro_espessura, width=12, font=font_subtitle,

                                state="readonly", justify="right")

entry_espessura_base.grid(row=0, column=2, pady=3, sticky="w")


tk.Label(frame_parametro_espessura, text="+/-", font=font_subtitle)\
    .grid(row=0, column=3, pady=3, sticky="nsew")


entry_tolerancia = tk.Entry(frame_parametro_espessura, width=12, font=font_subtitle,

                            state="readonly", justify="right")

entry_tolerancia.grid(row=0, column=4, pady=3, sticky="w")


tk.Label(frame_parametro_espessura, text="De:", font=font_subtitle)\
    .grid(row=1, column=1, sticky="e", pady=3)


entry_minimo = tk.Entry(frame_parametro_espessura, width=12, font=font_subtitle,

                        state="readonly", justify="right")

entry_minimo.grid(row=1, column=2, pady=3, sticky="w")


tk.Label(frame_parametro_espessura, text="à", font=font_subtitle)\
    .grid(row=1, column=3, pady=3, sticky="nsew")


entry_maximo = tk.Entry(frame_parametro_espessura, width=12, font=font_subtitle,

                        state="readonly", justify="right")

entry_maximo.grid(row=1, column=4, pady=3, sticky="w")


tk.Label(frame_parametro_espessura, text="Parâmetros de:", font=font_subtitle)\
    .grid(row=2, column=1, sticky="e", pady=3)


entry_datahora = tk.Entry(frame_parametro_espessura, width=25, font=font_subtitle,

                          state="readonly")

entry_datahora.grid(row=2, column=2, columnspan=3, pady=3, sticky="w")


def atualizar_parametro_espessura():

    if conn is None:

        for entry in [entry_espessura_base, entry_tolerancia, entry_minimo, entry_maximo, entry_datahora]:

            entry.config(state="normal")

            entry.delete(0, tk.END)

            entry.config(state="readonly")

        return

    cursor = conn.cursor()

    cursor.execute("SELECT minimo, maximo, datetime_col FROM range_emulsao")

    row = cursor.fetchone()

    if row:

        minimo, maximo, datetime_col = row.minimo, row.maximo, row.datetime_col


        # Converte para inteiro, arredondando os valores calculados

        valor_base = int(round((minimo + maximo) / 2))

        tolerancia = int(abs(valor_base - minimo))

        minimo_int = int(minimo)

        maximo_int = int(maximo)


        data_hora_str = datetime_col.strftime("%d/%m/%Y %H:%M:%S") if datetime_col else "sem data"


        # Preenche os campos com valores inteiros e texto da data/hora

        for entry, valor in [

            (entry_espessura_base, valor_base),

            (entry_tolerancia, tolerancia),

            (entry_minimo, minimo_int),

            (entry_maximo, maximo_int),

            (entry_datahora, data_hora_str)

        ]:

            entry.config(state="normal")

            entry.delete(0, tk.END)

            entry.insert(0, valor)

            entry.config(state="readonly")


    cursor.close()


atualizar_parametro_espessura()


# Linha em branco horizontal dentro do mesmo frame

tk.Label(frame_parametro_espessura, text="").grid(row=3, column=0, columnspan=7, pady=5)


label_espessura = tk.Label(frame_parametro_espessura, text="Espessura:", font=font_subtitle, width=12, anchor="e")

label_espessura.grid(row=4, column=0, padx=(0,2), pady=5, sticky="e")


label_espessura_1 = tk.Label(frame_parametro_espessura, text="Pós-esticagem (µm):", font=font_subtitle, width=15, anchor="e")

label_espessura_1.grid(row=4, column=1, padx=(0,2), pady=5, sticky="e")


entry_espessura_1 = tk.Entry(frame_parametro_espessura, width=8, font=font_subtitle, justify="right", state="readonly")

entry_espessura_1.grid(row=4, column=2, padx=(0,5), pady=5, sticky="w")


# Buscar valor do banco para pós-esticagem


label_espessura_2 = tk.Label(frame_parametro_espessura, text="Pós-Emulsão (µm):", font=font_subtitle, width=18, anchor="e")

label_espessura_2.grid(row=4, column=3, padx=(0,2), pady=5, sticky="e")


def validar_inteiro(new_value):

    if new_value == "":

        return True

    if new_value.startswith("-"):

        return False

    return new_value.isdigit()


vcmd = frame_parametro_espessura.register(validar_inteiro)


entry_espessura_2 = tk.Entry(

    frame_parametro_espessura, width=10, font=font_subtitle, justify="right",

    validate="key", validatecommand=(vcmd, "%P")

)

entry_espessura_2.grid(row=4, column=4, padx=(0,5), pady=5, sticky="w")


label_espessura_3 = tk.Label(frame_parametro_espessura, text="Emulsão (µm):", font=font_subtitle, width=13, anchor="e")

label_espessura_3.grid(row=4, column=5, padx=(0,2), pady=5, sticky="e")


entry_espessura_3 = tk.Entry(frame_parametro_espessura, width=10, font=font_subtitle, justify="right", state="readonly")

entry_espessura_3.grid(row=4, column=6, padx=(0,0), pady=5, sticky="w")


def atualizar_verificacao(event=None):

    try:

        emulsao_valor = float(entry_espessura_2.get())

    except ValueError:

        emulsao_valor = None


    try:

        pos_esticagem_valor_float = float(entry_espessura_1.get())

    except ValueError:

        pos_esticagem_valor_float = None


    entry_espessura_3.config(state="normal")

    entry_espessura_3.delete(0, tk.END)


    entry_status_emulsao.config(state="normal")

    entry_status_emulsao.delete(0, tk.END)


    if emulsao_valor is not None and pos_esticagem_valor_float is not None:

        diferenca = int(round(emulsao_valor - pos_esticagem_valor_float))


        # Puxa mínimo e máximo do banco

        try:

            cursor = conn.cursor()

            cursor.execute("SELECT TOP 1 minimo, maximo FROM range_emulsao ORDER BY id DESC")

            row = cursor.fetchone()

            cursor.close()

            if row:

                minimo_range, maximo_range = row

            else:

                minimo_range, maximo_range = 4, 8

        except Exception as e:

            print(f"Erro ao buscar range_emulsao: {e}")

            minimo_range, maximo_range = 4, 8


        entry_espessura_3.insert(0, str(diferenca))

        entry_espessura_3.config(state="readonly")


        # --- Atualiza status ---

        if diferenca < minimo_range or diferenca > maximo_range:

            entry_status_emulsao.insert(0, "NG")

            entry_status_emulsao.config(

                state="readonly",

                readonlybackground="red",

                fg="white"

            )

        else:

            entry_status_emulsao.insert(0, "OK")

            entry_status_emulsao.config(


                readonlybackground="white",

                fg="black"

            )

    else:

        # Campos vazios: limpar status e manter branco

        entry_espessura_3.config(state="readonly")

        entry_status_emulsao.config(

            state="readonly",

            readonlybackground="white",

            fg="black"

        )


# Vincular evento FocusOut

entry_espessura_2.bind("<FocusOut>", atualizar_verificacao)


func_emulsao_frame = tk.Frame(content_frame)

func_emulsao_frame.grid(row=13, column=0, columnspan=7, sticky="ew", padx=10, pady=5)


label_func_codigo_emulsao = tk.Label(func_emulsao_frame, text="Funcionário:", width=label_width, anchor="e", font=font_subtitle)

label_func_codigo_emulsao.grid(row=0, column=0, padx=5, pady=5, sticky="e")


entry_func_codigo_emulsao = tk.Entry(func_emulsao_frame, width=12, font=font_subtitle)

entry_func_codigo_emulsao.grid(row=0, column=1, padx=5, pady=5, sticky="w")


entry_func_nome_emulsao = tk.Entry(func_emulsao_frame, width=35, font=font_subtitle, state="readonly")

entry_func_nome_emulsao.grid(row=0, column=2, padx=5, pady=5, sticky="w")


label_data_emulsao = tk.Label(func_emulsao_frame, text="Data:", width=label_width, anchor="e", font=font_subtitle)

label_data_emulsao.grid(row=0, column=3, padx=5, pady=5, sticky="e")


entry_data_emulsao = tk.Entry(func_emulsao_frame, width=25, font=font_subtitle, state="readonly")

entry_data_emulsao.grid(row=0, column=4, padx=5, pady=5, sticky="w")


def on_matricula_emulsao_focusout(event):

    codigo_lido = entry_func_codigo_emulsao.get().strip()

    if not codigo_lido:

        # Limpa campos

        entry_func_nome_emulsao.configure(state="normal")

        entry_func_nome_emulsao.delete(0, tk.END)

        entry_func_nome_emulsao.configure(state="readonly")


        entry_data_emulsao.configure(state="normal")

        entry_data_emulsao.delete(0, tk.END)

        entry_data_emulsao.configure(state="readonly")

        return


    matricula = extrair_matricula(codigo_lido)

    if not matricula:

        messagebox.showwarning("Aviso", f"QR inválido: '{codigo_lido}'")

        entry_func_codigo_emulsao.delete(0, tk.END)


        entry_func_nome_emulsao.configure(state="normal")

        entry_func_nome_emulsao.delete(0, tk.END)

        entry_func_nome_emulsao.configure(state="readonly")


        entry_data_emulsao.configure(state="normal")

        entry_data_emulsao.delete(0, tk.END)

        entry_data_emulsao.configure(state="readonly")


        entry_func_codigo_emulsao.focus_set()

        return


    # AQUI, independente do status NG, aceita a matrícula e busca nome

    entry_func_codigo_emulsao.delete(0, tk.END)

    entry_func_codigo_emulsao.insert(0, matricula)


    nome = buscar_funcionario_por_matricula(matricula)

    if nome:

        entry_func_nome_emulsao.configure(state="normal")

        entry_func_nome_emulsao.delete(0, tk.END)

        entry_func_nome_emulsao.insert(0, nome)

        entry_func_nome_emulsao.configure(state="readonly")


        from datetime import datetime

        hoje = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        entry_data_emulsao.configure(state="normal")

        entry_data_emulsao.delete(0, tk.END)

        entry_data_emulsao.insert(0, hoje)

        entry_data_emulsao.configure(state="readonly")

    else:

        messagebox.showwarning("Aviso", f"Matrícula {matricula} não encontrada.")

        entry_func_codigo_emulsao.delete(0, tk.END)

        entry_func_nome_emulsao.configure(state="normal")

        entry_func_nome_emulsao.delete(0, tk.END)

        entry_func_nome_emulsao.configure(state="readonly")


        entry_data_emulsao.configure(state="normal")

        entry_data_emulsao.delete(0, tk.END)

        entry_data_emulsao.configure(state="readonly")


        entry_func_codigo_emulsao.focus_set()


entry_func_codigo_emulsao.bind("<FocusOut>", on_matricula_emulsao_focusout)


frame_apontamento_emulsao = tk.LabelFrame(content_frame, text="Apontamento - Emulsão", font=font_frame_title, padx=10, pady=10)

frame_apontamento_emulsao.grid(row=14, column=0, columnspan=7, sticky="ew", padx=10, pady=(10,20))


for col in range(7):

    frame_apontamento_emulsao.grid_columnconfigure(col, weight=1)


tk.Label(frame_apontamento_emulsao, text="Status:", font=font_subtitle, width=label_width, anchor="e").grid(row=0, column=0, padx=5, pady=5, sticky="e")

entry_status_emulsao = tk.Entry(frame_apontamento_emulsao, width=15, font=font_subtitle, justify="right")

entry_status_emulsao.grid(row=0, column=1, padx=5, pady=5, sticky="w")


tk.Label(frame_apontamento_emulsao, text="Observação:", font=font_subtitle, width=label_width, anchor="e").grid(row=1, column=0, padx=5, pady=5, sticky="ne")

entry_motivo_emulsao = tk.Text(frame_apontamento_emulsao, width=50, height=4, font=font_subtitle)

entry_motivo_emulsao.grid(row=1, column=1, columnspan=5, padx=5, pady=5, sticky="w")


#SALVAR EMULSÃO

def limpar_campos_esticagem():

    """Limpa todos os campos do bloco Esticagem"""

    # ID quadro (não limpa, pois é o principal)

    # combo origem e disposição

    combo_origem.set("")

    combo_origem.config(state="readonly")

    disp_var.set("")

    radio_normal.config(state="disabled")

    radio_bias.config(state="disabled")


    # Tensões

    for entry_x, entry_y in entries_tensao:

        entry_x.config(state="normal")

        entry_y.config(state="normal")

        entry_x.delete(0, tk.END)

        entry_y.delete(0, tk.END)

        entry_x.config(state="readonly")

        entry_y.config(state="readonly")


    # Funcionário

    entry_func_codigo.config(state="normal")

    entry_func_codigo.delete(0, tk.END)

    entry_func_codigo.config(state="readonly")

    entry_func_nome.config(state="normal")

    entry_func_nome.delete(0, tk.END)

    entry_func_nome.config(state="readonly")


    # Bias

    entry_bias.config(state="normal")

    entry_bias.delete(0, tk.END)

    entry_bias.config(state="disabled")


    # Mesh / Cola

    for entry_widget in [entry_mesh_leitura, entry_mesh_info, entry_cola_1, entry_cola_2]:

        entry_widget.config(state="normal")

        entry_widget.delete(0, tk.END)

        entry_widget.config(state="readonly")


    # Motivo, status e data

    entry_motivo.config(state="normal")

    entry_motivo.delete("1.0", tk.END)

    entry_motivo.config(state="disabled")

    entry_status.config(state="normal")

    entry_status.delete(0, tk.END)

    entry_status.config(state="readonly")

    entry_data.config(state="normal")

    entry_data.delete(0, tk.END)

    entry_data.config(state="readonly")


def limpar_campos_pos_esticagem_editavel():

    """Limpa todos os campos do bloco Pós-Esticagem e os deixa editáveis"""

    # Funcionário

    entry_func_codigo_pos.config(state="normal")

    entry_func_codigo_pos.delete(0, tk.END)


    entry_func_nome_pos.config(state="normal")

    entry_func_nome_pos.delete(0, tk.END)


    # Tensões 3D

    for entry_x, entry_y in entries_tensao_3d:

        entry_x.config(state="normal")

        entry_y.config(state="normal")

        entry_x.delete(0, tk.END)

        entry_y.delete(0, tk.END)


    # Média 3D

    entry_media_3d.config(state="normal")

    entry_media_3d.delete(0, tk.END)


    for entry in espessura_3d_entries:

        entry.config(state="normal")

        entry.delete(0, tk.END)


    # Motivo

    entry_motivo_pos.config(state="normal")

    entry_motivo_pos.delete("1.0", tk.END)


    # Status e data

    entry_status_pos.config(state="normal")

    entry_status_pos.delete(0, tk.END)


    entry_data_pos.config(state="normal")

    entry_data_pos.delete(0, tk.END)


def limpar_campos_pos_esticagem():

    """Limpa todos os campos do bloco Pós-Esticagem"""

    # Funcionário

    entry_func_codigo_pos.config(state="normal")

    entry_func_codigo_pos.delete(0, tk.END)

    entry_func_codigo_pos.config(state="readonly")


    entry_func_nome_pos.config(state="normal")

    entry_func_nome_pos.delete(0, tk.END)

    entry_func_nome_pos.config(state="readonly")


    # Tensões 3D

    for entry_x, entry_y in entries_tensao_3d:

        entry_x.config(state="normal")

        entry_y.config(state="normal")

        entry_x.delete(0, tk.END)

        entry_y.delete(0, tk.END)

        entry_x.config(state="readonly")

        entry_y.config(state="readonly")


    # Média 3D

    entry_media_3d.config(state="normal")

    entry_media_3d.delete(0, tk.END)

    entry_media_3d.config(state="readonly")


    # Motivo

    entry_motivo_pos.config(state="normal")

    entry_motivo_pos.delete("1.0", tk.END)

    entry_motivo_pos.config(state="disabled")


    # Status e data

    entry_status_pos.config(state="normal")

    entry_status_pos.delete(0, tk.END)

    entry_status_pos.config(state="readonly")


    entry_data_pos.config(state="normal")

    entry_data_pos.delete(0, tk.END)

    entry_data_pos.config(state="readonly")


def limpar_campos_emulsao():

    # --- Entrys inputáveis do usuário ---

    campos_para_limpar = [

        entry_emulsao_qr,

        entry_emulsao_qr_readonly,

        entry_espessura_1,

        entry_espessura_2,

        entry_espessura_3,

        entry_func_codigo_emulsao,

        entry_func_nome_emulsao,

        entry_data_emulsao,

        entry_status_emulsao

    ]


    for entry in campos_para_limpar:

        # Salva o estado atual

        estado_original = entry.cget("state")

        # Temporariamente permite apagar

        entry.config(state="normal")

        entry.delete(0, tk.END)

        # Restaura o estado original

        entry.config(state=estado_original)


    # --- Texts ---

    entry_motivo_emulsao.config(state="normal")

    entry_motivo_emulsao.delete("1.0", tk.END)

    entry_motivo_emulsao.config(state="normal")  # Text não tem readonly, mas bloqueado pode ser feito


    # --- Checkbuttons ---

    var_msfilm.set(False)


def limpar_campos_emulsao_editavel():

    """

    Limpa todos os campos de emulsão e mantém todos editáveis para o usuário.

    Inclui todos os campos, inclusive os que são apenas leitura na versão normal.

    """

    # --- Todos os Entrys relacionados à emulsão ---

    campos_para_limpar = [

        entry_emulsao_qr,

        entry_emulsao_qr_readonly,


        entry_espessura_2,

        entry_espessura_3,

        entry_func_codigo_emulsao,

        entry_func_nome_emulsao,

        entry_data_emulsao,

        entry_status_emulsao

    ]


    for entry in campos_para_limpar:

        entry.config(state="normal")  # garante que esteja editável

        entry.delete(0, tk.END)


    # --- Text editável ---

    entry_motivo_emulsao.config(state="normal")

    entry_motivo_emulsao.delete("1.0", tk.END)


    # --- Checkbuttons ---

    var_msfilm.set(False)


from datetime import datetime

import pyodbc

from tkinter import messagebox


def limpar_id_quadro():

    """Limpa o campo do ID do quadro e retorna o foco para ele."""

    entry_id_quadro.config(state="normal")

    entry_id_quadro.delete(0, tk.END)

    entry_id_quadro.focus_set()


def salvar_emulsao():

    try:

        cursor = conn.cursor()


        # --- Pegar ID do quadro ---

        id_quadro_val = entry_id_quadro.get().strip()

        if not id_quadro_val:

            messagebox.showwarning("Aviso", "Informe o ID do Quadro antes de salvar!")

            return


        # --- Verificar se o quadro existe em esticagem ---

        cursor.execute("SELECT COUNT(*) FROM dbo.esticagem WHERE id_quadro = ?", (id_quadro_val,))

        existe = cursor.fetchone()[0]

        if existe == 0:

            messagebox.showerror("Erro", "É necessário preencher Esticagem primeiro.")

            return


        # --- Verificar status da última emulsão ---

        cursor.execute("SELECT status FROM dbo.emulsao WHERE id_quadro = ? ORDER BY data DESC", (id_quadro_val,))

        row_status_emulsao = cursor.fetchone()

        emulsao_ok = row_status_emulsao and row_status_emulsao[0].upper() == "OK"


        # --- Verificar status da última revelação ---

        cursor.execute("SELECT status FROM dbo.revelacao WHERE id_quadro = ? ORDER BY data DESC", (id_quadro_val,))

        row_status_revelacao = cursor.fetchone()

        revelacao_ng = row_status_revelacao and row_status_revelacao[0].upper() == "NG"


        # --- Bloqueio de salvar emulsão ---

        if emulsao_ok and not revelacao_ng:

            messagebox.showwarning(

                "Aviso",

                "Já existe registro de emulsão com status OK e revelação não é NG. Não é permitido salvar novamente."

            )

            return


        # --- Obter demais dados da interface ---

        qr_emulsao_val = entry_emulsao_qr.get().strip()

        emulsao_val = entry_emulsao_qr_readonly.get().strip()

        polimero_val = var_msfilm.get()

        espessura_pos_emulsao_val = entry_espessura_2.get().strip()

        espessura_emulsao_val = entry_espessura_3.get().strip()

        from datetime import datetime

        data_val = datetime.now()

        reg_func_val = entry_func_codigo_emulsao.get().strip()

        status_val = entry_status_emulsao.get().strip()

        observacao_val = entry_motivo_emulsao.get("1.0", tk.END).strip()


        # --- Validações obrigatórias ---

        if not qr_emulsao_val or not emulsao_val or not espessura_pos_emulsao_val \
           or not espessura_emulsao_val or not reg_func_val or not status_val:

            messagebox.showwarning("Aviso", "Preencha todos os campos obrigatórios antes de salvar.")

            return


        # Observação obrigatória quando status for NG

        if status_val.upper() == "NG" and not observacao_val:

            messagebox.showwarning("Aviso", "Observação é obrigatória quando o status for NG.")

            entry_motivo_emulsao.focus_set()

            return


        # --- Inserir no banco ---

        cursor.execute("""

            INSERT INTO dbo.emulsao (

                id_quadro, qr_emulsao, emulsao, polimero,

                espessura_pos_emulsao, espessura_emulsao,

                data, reg_func, status, observacao

            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)

        """, (

            id_quadro_val, qr_emulsao_val, emulsao_val, polimero_val,

            espessura_pos_emulsao_val, espessura_emulsao_val,

            data_val, reg_func_val, status_val, observacao_val if observacao_val else None

        ))


        conn.commit()

        cursor.close()


        # --- Limpar apenas se salvou com sucesso ---

        limpar_campos_emulsao()

        limpar_campos_esticagem()

        limpar_campos_pos_esticagem()

        limpar_id_quadro()


        messagebox.showinfo("Sucesso", "Emulsão salva com sucesso!")


    except Exception as e:

        messagebox.showerror("Erro de Banco", f"Ocorreu um erro ao salvar a emulsão:\n{e}")


# --- Botão Salvar Emulsão ---

botao_salvar_emulsao = tk.Button(

    frame_apontamento_emulsao,

    text="Salvar Emulsão",

    font=font_subtitle,

    command=salvar_emulsao  # agora chama só a função principal

)

botao_salvar_emulsao.grid(row=2, column=0, columnspan=7, pady=(10,0))


titulo_revelacao = tk.Label(content_frame, text="4. Revelação", font=font_title)

titulo_revelacao.grid(row=15, column=0, columnspan=7, pady=(30, 25), sticky="nsew")


# ------------------- Frame Revelação -------------------

frame_revelacao = tk.Frame(content_frame)

frame_revelacao.grid(row=16, column=0, columnspan=7, sticky="ew", padx=10, pady=5)


# Configura 3 colunas principais do frame

for col in range(3):

    frame_revelacao.columnconfigure(col, weight=1)


# --- Labels e entries iniciais ---

label_fotolito = tk.Label(frame_revelacao, text="Fotolito (QR):", font=font_subtitle, anchor="e", width=15)

label_fotolito.grid(row=0, column=0, padx=5, pady=5, sticky="e")


entry_fotolito = tk.Entry(frame_revelacao, width=16, font=font_subtitle, justify="left")

entry_fotolito.grid(row=0, column=1, padx=5, pady=5, sticky="w")


entry_fotolito_readonly = tk.Entry(frame_revelacao, width=25, font=font_subtitle, justify="left", state="readonly")

entry_fotolito_readonly.grid(row=0, column=2, padx=5, pady=5, sticky="w")


entry_fotolito_extra = tk.Entry(frame_revelacao, width=40, font=font_subtitle, justify="left", state="readonly")

entry_fotolito_extra.grid(row=0, column=4, padx=5, pady=5, sticky="w")


label_modelo = tk.Label(frame_revelacao, text="Modelo:", font=font_subtitle, anchor="e", width=15)

label_modelo.grid(row=1, column=0, padx=5, pady=5, sticky="e")


entry_modelo_readonly = tk.Entry(frame_revelacao, width=25, font=font_subtitle, justify="left", state="readonly")

entry_modelo_readonly.grid(row=1, column=2, padx=5, pady=5, sticky="w")


label_apelido = tk.Label(frame_revelacao, text="Apelido (ALIAS):", font=font_subtitle, anchor="e", width=15)

label_apelido.grid(row=2, column=0, padx=5, pady=5, sticky="e")


entry_apelido_readonly = tk.Entry(frame_revelacao, width=25, font=font_subtitle, justify="left", state="readonly")

entry_apelido_readonly.grid(row=2, column=2, padx=5, pady=5, sticky="w")


# ---------------- Matriz 4x4 com cabeçalhos ----------------

frame_matriz = tk.Frame(frame_revelacao)

frame_matriz.grid(row=3, column=0, columnspan=7, sticky="nsew")  # abaixo de Modelo/Apelido


# Configura colunas da matriz (4 colunas agora)

for col in range(4):

    frame_matriz.columnconfigure(col, weight=1)


# --- Cabeçalhos ---

headers = ["Descrição", "Quadro", "Fotolito", "Verificação"]

for c, titulo in enumerate(headers):

    largura = 20 if titulo in ("Quadro", "Fotolito") else 12  # aumenta largura das duas colunas

    tk.Label(

        frame_matriz,

        text=titulo,

        font=font_subtitle,

        anchor="center",

        borderwidth=1,

        relief="solid",   # cabeçalho com borda mais forte

        width=largura

    ).grid(row=0, column=c, sticky="nsew")


# --- Descrições fixas ---

descricoes = ["Mesh", "Disposição", "Emulsão", "Uso/Destino"]


# --- Entradas da matriz (3 linhas x 3 colunas de dados), somente leitura ---

matriz_entries = []

for r, desc in enumerate(descricoes, start=1):

    linha_entries = []


    # Coluna de descrição (negrito, com borda leve)

    tk.Label(

        frame_matriz,

        text=desc,

        font=font_subtitle,   # mantém negrito

        anchor="center",

        borderwidth=1,

        relief="ridge",

        width=12

    ).grid(row=r, column=0, sticky="nsew")


    # Demais colunas (Quadro, Fotolito, Verificação)

    for c in range(1, 4):

        if c in (1, 2):   # Quadro e Fotolito mais largos

            largura = 20

        elif c == 3:      # Verificação menor

            largura = 12

        else:

            largura = 15


        e = tk.Entry(

            frame_matriz,

            width=largura,

            font=font_subtitle,

            justify="center",

            state="readonly",

            borderwidth=1,

            relief="ridge"

        )

        e.grid(row=r, column=c, sticky="nsew")

        linha_entries.append(e)


    matriz_entries.append(linha_entries)


# --- Função para validar QR Fotolito e preencher dados ---

def verificar_mesh(quadro_val, fotolito_val):

    """

    Retorna True se for considerado igual:

    - Comparação normal ignorando maiúsculas/minúsculas

    - Aceita Mesh menor (115) ser igual a maior correspondente (155)

    """

    q = quadro_val.upper().strip()

    f = fotolito_val.upper().strip()


    # Aceitar casos especiais

    if q.startswith("MESH 200 - 155") and f.startswith("MESH 200 - 115"):

        return True

    if q.startswith("MESH 255 - 155") and f.startswith("MESH 255 - 115"):

        return True

    # Comparação normal

    return q == f


def on_fotolito_enter(event=None):

    codigo = entry_fotolito.get().strip()


    if not codigo.endswith("F01"):

        messagebox.showerror("Erro", "Código Fotolito inválido!")

        entry_fotolito.delete(0, tk.END)

        return


    cod_fotolito = codigo[:-3].lstrip("0") or "0"

    id_quadro_val = entry_id_quadro.get().strip().upper()

    print(f"Função chamada. id_quadro='{id_quadro_val}', cod_fotolito='{cod_fotolito}'")


    try:

        with conn.cursor() as cursor:

            # --- Buscar dados do Fotolito (agora inclui Partlist) ---

            cursor.execute("""

                SELECT [Fotolito/tela], Modelo, Apelido, Partlist, Mesh, [Dispos.], Emulsão, [Uso/Destino]

                FROM fotolito_BD

                WHERE [Cód.Fotolito] = ?

            """, cod_fotolito)

            fotolito_row = cursor.fetchone()


            # --- Buscar dados da tabela esticagem ---

            cursor.execute("""

                SELECT mesh, disposicao, qr_mesh, origem

                FROM esticagem

                WHERE UPPER(id_quadro) = ?

            """, id_quadro_val)

            esticagem_row = cursor.fetchone()


            # --- Buscar dados da tabela emulsao ---

            cursor.execute("""

                SELECT emulsao, qr_emulsao

                FROM emulsao

                WHERE UPPER(id_quadro) = ?

            """, id_quadro_val)

            emulsao_row = cursor.fetchone()


        # --- Preencher campos read-only Fotolito ---

        if fotolito_row:

            for entry_widget, value in zip(

                [entry_fotolito_readonly, entry_modelo_readonly, entry_apelido_readonly, entry_fotolito_extra],

                fotolito_row[:4]  # agora pega as 4 primeiras colunas

            ):

                entry_widget.config(state="normal")

                entry_widget.delete(0, tk.END)

                entry_widget.insert(0, value or "")

                entry_widget.config(state="readonly")

        else:

            messagebox.showwarning("Aviso", f"Nenhum registro encontrado para código {cod_fotolito}")

            entry_fotolito.delete(0, tk.END)

            return


        # --- Preparar valores para a matriz ---

        valores_fotolito = [

            (fotolito_row[4] or "").strip(),  # Mesh

            (fotolito_row[5] or "").strip(),  # Disposição

            (fotolito_row[6] or "").strip(),  # Emulsão

            (fotolito_row[7] or "").strip()   # Uso/Destino Fotolito

        ]


        valores_quadro = [

            (esticagem_row[0] or "Não encontrado").strip() if esticagem_row else "Não encontrado",

            (esticagem_row[1] or "Não encontrado").strip() if esticagem_row else "Não encontrado",

            (emulsao_row[0] or "Não encontrado").strip() if emulsao_row else "Não encontrado"

        ]


        # --- Ajuste de exibição do Uso/Destino Quadro ---

        if esticagem_row:

            origem_val = (esticagem_row[3] or "").strip().upper()

            if origem_val == "NACIONAL":

                valores_quadro.append("Tinta/Gloss/Matt")

            elif origem_val == "IMPORTADO":

                valores_quadro.append("Gloss/Matt")

            else:

                valores_quadro.append(origem_val or "Não encontrado")

        else:

            valores_quadro.append("Não encontrado")


        # --- Verificação ---

        valores_verificacao = [

            "OK" if verificar_mesh(valores_quadro[0], valores_fotolito[0]) else "NG",

            "OK" if valores_quadro[1].upper() == valores_fotolito[1].upper() else "NG",

            "OK" if valores_quadro[2].upper() == valores_fotolito[2].upper() else "NG",

            # Uso/Destino: se origem Importado, fotolito deve ser Gloss/Matt

            "OK" if not (origem_val == "IMPORTADO" and valores_fotolito[3].upper() != "GLOSS/MATT") else "NG"

        ]


        # --- Preencher a matriz 4x4 ---

        for i in range(4):

            e_quadro = matriz_entries[i][0]

            e_quadro.config(state="normal")

            e_quadro.delete(0, tk.END)

            e_quadro.insert(0, valores_quadro[i])

            e_quadro.config(state="readonly")


            e_fotolito = matriz_entries[i][1]

            e_fotolito.config(state="normal")

            e_fotolito.delete(0, tk.END)

            e_fotolito.insert(0, valores_fotolito[i])

            e_fotolito.config(state="readonly")


            e_verif = matriz_entries[i][2]

            e_verif.config(state="normal")

            e_verif.delete(0, tk.END)

            e_verif.insert(0, valores_verificacao[i])

            e_verif.config(readonlybackground="red" if valores_verificacao[i] == "NG" else "white")

            e_verif.config(state="readonly")


        # --- Atualiza o Entry de Status geral ---

        status_geral = "NG" if "NG" in valores_verificacao[:4] else "OK"

        entry_status_revelacao.config(state="normal")

        entry_status_revelacao.delete(0, tk.END)

        entry_status_revelacao.insert(0, status_geral)

        entry_status_revelacao.config(readonlybackground="red" if status_geral == "NG" else "white")

        entry_status_revelacao.config(state="readonly")


        entry_func_codigo_revelacao.focus_set()


    except Exception as e:

        messagebox.showerror("Erro", str(e))

        entry_fotolito.delete(0, tk.END)

        entry_fotolito.focus_set()


# --- Bind para Enter e FocusOut ---

entry_fotolito.bind("<Return>", lambda event: on_fotolito_enter())

entry_fotolito.bind("<FocusOut>", on_fotolito_enter)


# ---------------- Frame funcionário / data

func_revelacao_frame = tk.Frame(frame_revelacao)

func_revelacao_frame.grid(row=4, column=0, columnspan=7, sticky="ew", padx=10, pady=(15, 5))


label_func_codigo_revelacao = tk.Label(func_revelacao_frame, text="Funcionário:", width=label_width, anchor="e", font=font_subtitle)

label_func_codigo_revelacao.grid(row=0, column=0, padx=5, pady=5, sticky="e")


entry_func_codigo_revelacao = tk.Entry(func_revelacao_frame, width=12, font=font_subtitle)

entry_func_codigo_revelacao.grid(row=0, column=1, padx=5, pady=5, sticky="w")


entry_func_nome_revelacao = tk.Entry(func_revelacao_frame, width=35, font=font_subtitle, state="readonly")

entry_func_nome_revelacao.grid(row=0, column=2, padx=5, pady=5, sticky="w")


label_data_revelacao = tk.Label(func_revelacao_frame, text="Data:", width=label_width, anchor="e", font=font_subtitle)

label_data_revelacao.grid(row=0, column=3, padx=5, pady=5, sticky="e")


entry_data_revelacao = tk.Entry(func_revelacao_frame, width=25, font=font_subtitle, state="readonly")

entry_data_revelacao.grid(row=0, column=4, padx=5, pady=5, sticky="w")


# Função de leitura QR / matrícula

def on_matricula_revelacao_enter(event=None):

    codigo_lido = entry_func_codigo_revelacao.get().strip()

    if not codigo_lido:

        entry_func_nome_revelacao.configure(state="normal")

        entry_func_nome_revelacao.delete(0, tk.END)

        entry_func_nome_revelacao.configure(state="readonly")


        entry_data_revelacao.configure(state="normal")

        entry_data_revelacao.delete(0, tk.END)

        entry_data_revelacao.configure(state="readonly")

        entry_status_revelacao.focus_set()

        return


    matricula = extrair_matricula(codigo_lido)

    if not matricula:

        messagebox.showwarning("Aviso", f"QR inválido: '{codigo_lido}'")

        entry_func_codigo_revelacao.delete(0, tk.END)

        entry_func_codigo_revelacao.focus_set()

        return


    entry_func_codigo_revelacao.delete(0, tk.END)

    entry_func_codigo_revelacao.insert(0, matricula)


    nome = buscar_funcionario_por_matricula(matricula)

    if nome:

        entry_func_nome_revelacao.configure(state="normal")

        entry_func_nome_revelacao.delete(0, tk.END)

        entry_func_nome_revelacao.insert(0, nome)

        entry_func_nome_revelacao.configure(state="readonly")


        from datetime import datetime

        hoje = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        entry_data_revelacao.configure(state="normal")

        entry_data_revelacao.delete(0, tk.END)

        entry_data_revelacao.insert(0, hoje)

        entry_data_revelacao.configure(state="readonly")


        entry_status_revelacao.focus_set()

    else:

        messagebox.showwarning("Aviso", f"Matrícula {matricula} não encontrada.")

        entry_func_codigo_revelacao.delete(0, tk.END)

        entry_func_codigo_revelacao.focus_set()


entry_func_codigo_revelacao.bind("<Return>", on_matricula_revelacao_enter)


# Apontamento - Revelação

frame_apontamento_revelacao = tk.LabelFrame(content_frame, text="Apontamento - Revelação (Inicial)                                               Apontamento - Revelação (Final)", font=font_frame_title, padx=10, pady=10)

frame_apontamento_revelacao.grid(row=17, column=0, columnspan=7, sticky="ew", padx=10, pady=(10, 20))


frame_apontamento_revelacao.grid_columnconfigure(0, weight=1)

frame_apontamento_revelacao.grid_columnconfigure(1, weight=1)

frame_apontamento_revelacao.grid_rowconfigure(99, weight=1)


# Frame esquerdo

frame_apontamento_esquerda = tk.Frame(frame_apontamento_revelacao)

frame_apontamento_esquerda.grid(row=0, column=0, sticky="nsew", padx=(0, 5))


# --- Status ---

tk.Label(frame_apontamento_esquerda, text="Status:", font=font_subtitle, width=label_width, anchor="e")\
    .grid(row=0, column=0, padx=0, pady=5, sticky="w")  # label encostada

entry_status_revelacao = tk.Entry(frame_apontamento_esquerda, width=15, font=font_subtitle, justify="right")

entry_status_revelacao.grid(row=0, column=1, padx=0, pady=5, sticky="w")  # entry encostada


# --- Observação ---

tk.Label(frame_apontamento_esquerda, text="Observação:", font=font_subtitle, width=label_width, anchor="e")\
    .grid(row=1, column=0, padx=0, pady=5, sticky="nw")  # label encostada

entry_motivo_revelacao = tk.Text(frame_apontamento_esquerda, width=30, height=5, font=font_subtitle)

entry_motivo_revelacao.grid(row=1, column=1, padx=0, pady=5, sticky="w")  # texto encosta


def limpar_campos_revelacao():

    """

    Limpa todos os campos inputáveis do módulo de revelação,

    sem alterar o state atual dos widgets (readonly ou normal).

    """

    # --- Campos Entry ---

    campos_entry = [

        entry_fotolito,

        entry_func_codigo_revelacao,

        entry_status_revelacao

    ]


    for entry in campos_entry:

        entry_state = entry.cget("state")  # guarda o state atual

        entry.config(state="normal")       # temporariamente habilita para apagar

        entry.delete(0, tk.END)

        entry.config(state=entry_state)    # volta ao state original


    # --- Campos Text ---

    campos_text = [

        entry_motivo_revelacao

    ]

    for text_widget in campos_text:

        text_widget.delete("1.0", tk.END)


    # --- Limpa a matriz 4x4 ---

    for linha in matriz_entries:

        for e in linha:

            e_state = e.cget("state")

            e.config(state="normal")

            e.delete(0, tk.END)

            e.config(state=e_state)


    # --- Campos readonly extras (fotolito, modelo, apelido, extra) ---

    extras = [

        entry_fotolito_readonly,

        entry_fotolito_extra,

        entry_modelo_readonly,

        entry_apelido_readonly,

        entry_func_nome_revelacao,

        entry_data_revelacao

    ]

    for e in extras:

        e_state = e.cget("state")

        e.config(state="normal")

        e.delete(0, tk.END)

        e.config(state=e_state)


def salvar_revelacao():

    status_val = entry_status_revelacao.get().strip().upper()


    # --- Verifica se o status atual é NG ---

    if status_val == "NG":

        messagebox.showwarning("Atenção", "Verificar Quadro ou Fotolito antes de salvar!")

        return


    # --- Verifica se já existe registro OK para este quadro ---

    id_quadro_val = entry_id_quadro.get().strip().upper()

    if not id_quadro_val:

        messagebox.showwarning("Aviso", "Informe o ID do Quadro antes de salvar!")

        return


    try:

        cursor = conn.cursor()

        cursor.execute("""

            SELECT TOP 1 status

            FROM revelacao

            WHERE id_quadro = ?

            ORDER BY data DESC

        """, (id_quadro_val,))

        row = cursor.fetchone()

        cursor.close()


        if row and row[0].strip().upper() == "OK":

            messagebox.showwarning(

                "Aviso",

                "Já existe registro de Revelação com status OK para este quadro. Não é permitido salvar novamente."

            )

            return


        # --- Valores a serem salvos ---

        qr_fotolito_val = entry_fotolito.get().strip()

        fotolito_val = entry_fotolito_extra.get().strip()

        modelo_val = entry_modelo_readonly.get().strip()

        apelido_val = entry_apelido_readonly.get().strip()

        reg_func_val = entry_func_codigo_revelacao.get().strip()

        from datetime import datetime

        data_val = datetime.now()

        observacao_val = entry_motivo_revelacao.get("1.0", tk.END).strip()


        # --- Inserção no banco ---

        with conn.cursor() as cursor:

            cursor.execute("""

                INSERT INTO revelacao (

                    qr_fotolito,

                    fotolito,

                    modelo,

                    apelido,

                    reg_func,

                    data,

                    status,

                    observacao,

                    id_quadro

                )

                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)

            """, qr_fotolito_val, fotolito_val, modelo_val, apelido_val,

                 reg_func_val, data_val, status_val, observacao_val, id_quadro_val)

            conn.commit()


        messagebox.showinfo("Sucesso", "Revelação salva com sucesso!")


        # --- Limpa todos os campos inputáveis ---

        limpar_campos_revelacao()

        limpar_campos_esticagem()

        limpar_campos_emulsao_editavel()

        limpar_campos_pos_esticagem_editavel()

        limpar_id_quadro()


    except Exception as e:

        messagebox.showerror("Erro", f"Falha ao salvar revelação:\n{str(e)}")


def limpar_campos_revelacao_editavel():

    """

    Limpa todos os campos inputáveis do módulo de revelação.

    - Campos editáveis permanecem editáveis (state="normal").

    - Campos somente leitura permanecem readonly.

    """


    # --- Campos Entry (editáveis pelo usuário) ---

    campos_entry_editaveis = [

        entry_fotolito,

        entry_func_codigo_revelacao,

        entry_status_revelacao

    ]

    for entry in campos_entry_editaveis:

        entry.config(state="normal")

        entry.delete(0, tk.END)


    # --- Campo Text (observação é inputável) ---

    entry_motivo_revelacao.config(state="normal")

    entry_motivo_revelacao.delete("1.0", tk.END)


    # --- Matriz 4x4 (editáveis) ---

    for linha in matriz_entries:

        for e in linha:

            e.config(state="normal")

            e.delete(0, tk.END)


    # --- Campos somente leitura (devem continuar readonly) ---

    campos_readonly = [

        entry_fotolito_readonly,

        entry_fotolito_extra,

        entry_modelo_readonly,

        entry_apelido_readonly,

        entry_func_nome_revelacao,

        entry_data_revelacao

    ]

    for e in campos_readonly:

        e.config(state="normal")   # libera só para apagar

        e.delete(0, tk.END)

        e.config(state="readonly") # volta readonly


# Configura o grid do frame para expandir e empurrar o botão

frame_apontamento_esquerda.rowconfigure(99, weight=1)  # Linha "vazia" que ocupa o espaço

frame_apontamento_esquerda.columnconfigure(0, weight=1)


botao_salvar_revelacao = tk.Button(

    frame_apontamento_esquerda,

    text="Salvar Revelação - Início",

    font=font_subtitle,

    command=salvar_revelacao

)

botao_salvar_revelacao.grid(

    row=100, column=0, columnspan=7,

    pady=(10, 0),

    sticky="s"   # Cola no fundo

)


separator = tk.Frame(frame_apontamento_revelacao, width=2, bg="gray")

separator.grid(row=0, column=1, sticky="ns", pady=5)


# Frame direito

frame_apontamento_direita = tk.Frame(frame_apontamento_revelacao)

frame_apontamento_direita.grid(row=0, column=2, sticky="nsew", padx=(5, 0))


# --- Frame para Fotolito (QR) ---

frame_fotolito = tk.Frame(frame_apontamento_direita)

frame_fotolito.grid(row=0, column=0, columnspan=3, sticky="w", padx=5, pady=5)


tk.Label(frame_fotolito, text="Fotolito (QR):", font=font_subtitle, anchor="w").pack(side="left")

entry_fotolito_final = tk.Entry(frame_fotolito, width=25, font=font_subtitle, justify="left")

entry_fotolito_final.pack(side="left", padx=(2,20))


# --- Frame para Status e Reutilizar quadro ---

frame_status_reutilizar = tk.Frame(frame_apontamento_direita)

frame_status_reutilizar.grid(row=1, column=0, columnspan=3, sticky="w", padx=5, pady=5)


# --- Status ---

tk.Label(frame_status_reutilizar, text="Status:", font=font_subtitle, anchor="w").pack(side="left")

entry_status_revelacao_final = tk.Entry(frame_status_reutilizar, width=15, font=font_subtitle, justify="right")

entry_status_revelacao_final.pack(side="left", padx=(2,20))


# --- Reutilizar Quadro ---

tk.Label(frame_status_reutilizar, text="Reutilizar quadro:", font=font_subtitle, anchor="w").pack(side="left")

var_reutilizar_quadro = tk.StringVar(value="")  # nenhum selecionado por padrão


radio_sim = tk.Radiobutton(frame_status_reutilizar, text="Sim", variable=var_reutilizar_quadro,

                           value="Sim", font=font_subtitle, state="disabled")

radio_sim.pack(side="left", padx=(5,2))


radio_nao = tk.Radiobutton(frame_status_reutilizar, text="Não", variable=var_reutilizar_quadro,

                           value="Não", font=font_subtitle, state="disabled")

radio_nao.pack(side="left", padx=(2,0))


# --- Caixa de Observação ---

tk.Label(frame_apontamento_direita, text="Observação:", font=font_subtitle, anchor="w").grid(

    row=2, column=0, padx=5, pady=(5,2), sticky="ne"

)

entry_observacao = tk.Text(frame_apontamento_direita, width=40, height=4, font=font_subtitle)

entry_observacao.grid(row=2, column=1, columnspan=2, padx=5, pady=(5,2), sticky="w")


def on_fotolito_final_enter(event=None):

    valor_input = entry_fotolito_final.get().strip()

    valor_registro = entry_fotolito.get().strip()


    # --- Atualiza status e aparência ---

    entry_status_revelacao_final.config(state="normal")  # habilita temporariamente


    if valor_input == valor_registro:

        # Status OK → editável, fundo branco

        entry_status_revelacao_final.delete(0, tk.END)

        entry_status_revelacao_final.insert(0, "OK")

        entry_status_revelacao_final.config(readonlybackground="white", foreground="black")

        entry_status_revelacao_final.config(state="normal")

    else:

        # Status NG → readonly, fundo vermelho, letra branca

        entry_status_revelacao_final.delete(0, tk.END)

        entry_status_revelacao_final.insert(0, "NG")

        entry_status_revelacao_final.config(readonlybackground="red", foreground="white")

        entry_status_revelacao_final.config(state="readonly")


    # --- Habilita ou desabilita os rádios de Reutilizar Quadro ---

    status = entry_status_revelacao_final.get().strip().upper()

    if status == "NG":

        radio_sim.config(state="normal")

        radio_nao.config(state="normal")

    else:

        var_reutilizar_quadro.set("")  # limpa seleção, nenhum marcado

        radio_sim.config(state="disabled")

        radio_nao.config(state="disabled")


    # --- Pula automaticamente para o campo de Observação ---

    entry_observacao.focus_set()


# Bind para Enter

entry_fotolito_final.bind("<Return>", on_fotolito_final_enter)


# --- Funcionário (QR e Nome) ---

tk.Label(frame_apontamento_direita, text="Funcionário:", font=font_subtitle, anchor="w").grid(

    row=3, column=0, padx=5, pady=(5,2), sticky="w"

)

entry_funcionario_qr_final = tk.Entry(frame_apontamento_direita, width=10, font=font_subtitle, justify="left")

entry_funcionario_qr_final.grid(row=3, column=1, padx=5, pady=(5,2), sticky="w")


entry_funcionario_nome_final = tk.Entry(frame_apontamento_direita, width=30, font=font_subtitle, state="readonly")

entry_funcionario_nome_final.grid(row=3, column=2, padx=5, pady=(5,2), sticky="w")


def on_matricula_revelacao_final_enter(event=None):

    codigo_lido = entry_funcionario_qr_final.get().strip()

    if not codigo_lido:

        entry_funcionario_nome_final.configure(state="normal")

        entry_funcionario_nome_final.delete(0, tk.END)

        entry_funcionario_nome_final.configure(state="readonly")

        return


    matricula = extrair_matricula(codigo_lido)

    if not matricula:

        messagebox.showwarning("Aviso", f"QR inválido: '{codigo_lido}'")

        entry_funcionario_qr_final.delete(0, tk.END)

        entry_funcionario_qr_final.focus_set()

        return


    entry_funcionario_qr_final.delete(0, tk.END)

    entry_funcionario_qr_final.insert(0, matricula)


    nome = buscar_funcionario_por_matricula(matricula)

    if nome:

        entry_funcionario_nome_final.configure(state="normal")

        entry_funcionario_nome_final.delete(0, tk.END)

        entry_funcionario_nome_final.insert(0, nome)

        entry_funcionario_nome_final.configure(state="readonly")

    else:

        messagebox.showwarning("Aviso", f"Matrícula {matricula} não encontrada.")

        entry_funcionario_qr_final.delete(0, tk.END)

        entry_funcionario_qr_final.focus_set()


# Bind para capturar Enter no campo de QR

entry_funcionario_qr_final.bind("<Return>", on_matricula_revelacao_final_enter)


def salvar_revelacao_final():

    try:

        id_quadro = entry_id_quadro.get().strip()

        qr_fotolito = entry_fotolito_final.get().strip()

        reg_func = entry_funcionario_qr_final.get().strip()

        status = entry_status_revelacao_final.get().strip().upper()

        observacao = entry_observacao.get("1.0", tk.END).strip()

        reutilizar = var_reutilizar_quadro.get()  # Sim ou Não


        # --- Verifica se o processo anterior (revelação) possui registro ---

        cursor = conn.cursor()

        cursor.execute("SELECT COUNT(*) FROM revelacao WHERE id_quadro = ?", (id_quadro,))

        registro_anterior = cursor.fetchone()[0]

        if registro_anterior == 0:

            messagebox.showwarning("Aviso", "Não é possível salvar a Revelação Final. Processo anterior não possui registro.")

            return


        # --- Validação obrigatória de campos ---

        if not id_quadro or not qr_fotolito or not reg_func or not status:

            messagebox.showwarning("Aviso", "Preencha todos os campos obrigatórios antes de salvar.")

            return


        # --- Se status NG, observação e checkbox obrigatório ---

        if status == "NG":

            if not observacao:

                messagebox.showwarning("Aviso", "Observação é obrigatória quando o status for NG.")

                entry_observacao.focus_set()

                return

            if reutilizar not in ("Sim", "Não"):

                messagebox.showwarning("Aviso", "Selecione uma opção para 'Reutilizar Quadro' quando o status for NG.")

                return


        # --- Se status OK, checkbox não é obrigatório ---

        if status == "OK":

            reutilizar = reutilizar or "N/D"


        # --- Data atual ---

        from datetime import datetime

        data = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


        # --- Inserção no banco ---

        cursor.execute("""

            INSERT INTO revelacao_final (id_quadro, qr_fotolito, reg_func, status, observacao, reutilizar, data)

            VALUES (?, ?, ?, ?, ?, ?, ?)

        """, (id_quadro, qr_fotolito, reg_func, status, observacao, reutilizar, data))

        conn.commit()


        messagebox.showinfo("Sucesso", "Revelação final salva com sucesso!")


        # --- Limpeza dos campos ---

        entry_fotolito_final.delete(0, tk.END)

        entry_status_revelacao_final.config(state="normal")

        entry_status_revelacao_final.delete(0, tk.END)

        entry_status_revelacao_final.config(state="readonly")

        entry_observacao.delete("1.0", tk.END)

        entry_funcionario_qr_final.delete(0, tk.END)

        entry_funcionario_nome_final.config(state="normal")

        entry_funcionario_nome_final.delete(0, tk.END)

        entry_funcionario_nome_final.config(state="readonly")

        var_reutilizar_quadro.set("")

        radio_sim.config(state="disabled")

        radio_nao.config(state="disabled")

        limpar_campos_emulsao_editavel()

        limpar_campos_esticagem()

        limpar_campos_revelacao_editavel()

        limpar_campos_pos_esticagem_editavel()

        limpar_id_quadro()


    except Exception as e:

        messagebox.showerror("Erro", f"Erro ao salvar revelação final: {e}")


# Botão para salvar no banco

botao_salvar_revelacao = tk.Button(

    frame_apontamento_direita,

    text="Salvar Revelação - Final",

    font=font_subtitle,

    width=20,

    command=salvar_revelacao_final

)

botao_salvar_revelacao.grid(row=4, column=0, columnspan=3, padx=5, pady=(10,5))


def verificar_status_revelacao(event=None):

    """

    Função que verifica o status da revelação final.

    - Se NG, habilita os rádios Sim/Não.

    - Se OK, desabilita os rádios e limpa a seleção.

    """

    status = entry_status_revelacao_final.get().strip().upper()


    if status == "NG":

        radio_sim.config(state="normal")

        radio_nao.config(state="normal")

    elif status == "OK":

        var_reutilizar_quadro.set("")  # limpa seleção

        radio_sim.config(state="disabled")

        radio_nao.config(state="disabled")


# ou para atualizar imediatamente ao digitar

entry_status_revelacao_final.bind("<KeyRelease>", verificar_status_revelacao)


def carregar_revelacao_por_quadro(id_quadro):

    """

    Carrega os dados de revelação com base no id_quadro.

    - Se houver registro em revelacao_final com NG + Sim,

      só permite carregar revelação se a data da revelação for mais recente.

    - Campos essenciais: QR Fotolito, funcionário, status, observação, nome do funcionário e data.

    - Campos complementares (modelo, apelido, uso/destino, matriz 4x4) são preenchidos

      automaticamente com base no QR do fotolito.

    - Se houver dados, todos os campos se tornam readonly.

    - Se não houver registro, não altera nada.

    """

    if not id_quadro or not id_quadro.strip():

        return  # não faz nada se id_quadro vazio


    id_quadro = id_quadro.strip().upper()


    try:

        # --- 1. Buscar último registro de revelacao_final ---

        with conn.cursor() as cursor:

            cursor.execute("""

                SELECT TOP 1 status, reutilizar, data

                FROM revelacao_final

                WHERE UPPER(id_quadro) = ?

                ORDER BY data DESC

            """, id_quadro)

            row_rev = cursor.fetchone()


        status_rev = reutilizar_rev = data_rev = None

        if row_rev:

            status_rev, reutilizar_rev, data_rev = row_rev


        # --- 2. Buscar último registro de revelação ---

        with conn.cursor() as cursor:

            cursor.execute("""

                SELECT qr_fotolito, reg_func, status, observacao, data

                FROM revelacao

                WHERE UPPER(id_quadro) = ?

                ORDER BY data DESC

            """, id_quadro)

            row = cursor.fetchone()


        if not row:

            limpar_campos_revelacao_editavel()

            return  # não faz nada se não houver registro


        qr_fotolito, reg_func, status, observacao, data_reg = row


        # --- 3. Bloqueio se revelacao_final = NG + Sim ---

        if status_rev and status_rev.upper() == "NG" and reutilizar_rev == "Sim":

            if not data_reg or (data_rev and data_reg <= data_rev):

                # Revelação não foi refeita → não carregar

                limpar_campos_revelacao_editavel()

                print("DEBUG: Revelação final NG + Sim e revelação não refeita → bloqueado.")

                return

            else:

                print("DEBUG: Revelação refeita após NG + Sim → carregando normalmente.")


        # --- 4. Preencher campos essenciais ---

        # QR Fotolito

        entry_fotolito.config(state="normal")

        entry_fotolito.delete(0, tk.END)

        entry_fotolito.insert(0, qr_fotolito or "")

        entry_fotolito.config(state="readonly")


        # Código do funcionário

        entry_func_codigo_revelacao.config(state="normal")

        entry_func_codigo_revelacao.delete(0, tk.END)

        entry_func_codigo_revelacao.insert(0, str(reg_func or ""))

        entry_func_codigo_revelacao.config(state="disabled")


        # Nome do funcionário

        nome_func = buscar_funcionario_por_matricula(reg_func) if reg_func else ""

        entry_func_nome_revelacao.config(state="normal")

        entry_func_nome_revelacao.delete(0, tk.END)

        entry_func_nome_revelacao.insert(0, nome_func or "")

        entry_func_nome_revelacao.config(state="readonly")


        # Data

        entry_data_revelacao.config(state="normal")

        entry_data_revelacao.delete(0, tk.END)

        entry_data_revelacao.insert(0, data_reg.strftime("%d/%m/%Y %H:%M:%S") if data_reg else "")

        entry_data_revelacao.config(state="readonly")


        # Status

        entry_status_revelacao.config(state="normal")

        entry_status_revelacao.delete(0, tk.END)

        entry_status_revelacao.insert(0, status or "")

        entry_status_revelacao.config(state="readonly")


        # Observação

        entry_motivo_revelacao.config(state="normal")

        entry_motivo_revelacao.delete("1.0", tk.END)

        entry_motivo_revelacao.insert("1.0", observacao or "")

        entry_motivo_revelacao.config(state="disabled")


        # --- 5. Preencher campos complementares via função existente ---

        if qr_fotolito:

            on_fotolito_enter()


        print("DEBUG: Dados da revelação carregados com sucesso.")


    except Exception as e:

        messagebox.showerror("Erro", f"Falha ao carregar dados da revelação:\n{e}")


combo_origem.grid_position = {

    "row": 1,

    "column": 3,

    "sticky": "w",

    "padx": 5,

    "pady": 5

}


# Frame final dos botões


button_frame_final = tk.Frame(content_frame)

button_frame_final.grid(row=18, column=0, columnspan=7, pady=(20, 10))


btn_sair_final = tk.Button(button_frame_final, text="Sair", width=15, command=root.quit)

btn_sair_final.pack(side="left", padx=10)


entries_emulsao_to_bind = [

    entry_emulsao_qr,

    entry_emulsao_qr_readonly,


    entry_espessura_1,

    entry_espessura_2,

    entry_espessura_3,

    entry_func_codigo_emulsao,

    entry_func_nome_emulsao,

    entry_data_emulsao,

    entry_status_emulsao,

    entry_motivo_emulsao,

]


entries_emulsao_editaveis = [

    entry_emulsao_qr,

    entry_espessura_2,

    entry_func_codigo_emulsao,

    entry_status_emulsao,

    entry_motivo_emulsao,

]


def focus_next_widget(event):

    event.widget.tk_focusNext().focus()

    return "break"


entries_to_bind = [

    entry_id_quadro,

    entry_mesh_leitura,

    entry_bias,

    entry_cola_1,

    entry_tensao,

    entry_tensao_plusminus,

    entry_de,

    entry_ate,

    entry_parametros,

    entry_func_codigo,

    entry_func_nome,

    entry_data,

]


for entry_x, entry_y in entries_tensao:

    entries_to_bind.append(entry_x)

    entries_to_bind.append(entry_y)


entries_to_bind += [

    entry_tensao_3d,

    entry_tensao_plusminus_3d,

    entry_de_3d,

    entry_ate_3d,

    entry_parametros_3d,

    *espessura_3d_entries,

    entry_func_codigo_pos,

    entry_func_nome_pos,

    entry_data_pos,

]


for entry_x, entry_y in entries_tensao_3d:

    entries_to_bind.append(entry_x)

    entries_to_bind.append(entry_y)


entries_to_bind.append(entry_status)

entries_to_bind.append(entry_status_pos)


for entry in entries_to_bind:

    if isinstance(entry, tk.Entry):

        entry.bind("<Return>", focus_next_widget)


def atualizar_cor_status(entry):

    texto = entry.get().strip().upper()

    if texto == "NG":

        entry.config(bg="red")

    else:

        entry.config(bg="white")


entry_status.bind("<KeyRelease>", lambda e: atualizar_cor_status(entry_status))

entry_status_pos.bind("<KeyRelease>", lambda e: atualizar_cor_status(entry_status_pos))


def on_enter_emulsao(event):

    widget = event.widget


    if widget == entry_espessura_2:

        atualizar_verificacao()


    try:

        idx = entries_emulsao_editaveis.index(widget)

        next_idx = (idx + 1) % len(entries_emulsao_editaveis)

        entries_emulsao_editaveis[next_idx].focus_set()

    except ValueError:

        pass


    if isinstance(widget, tk.Entry):

        return "break"


for e in entries_emulsao_editaveis:

    e.bind("<Return>", on_enter_emulsao)


def on_ctrl_enter_motivo(event):

    entry_status_emulsao.focus_set()

    return "break"


entry_motivo_emulsao.bind("<Control-Return>", on_ctrl_enter_motivo)


def buscar_emulsao_por_qr(qr):

    """

    Extrai o número de referência, verifica sufixo 'MPI' e se Tipo = 'Emulsão'.

    Retorna o valor da coluna Desc_Proc_01 da tabela mp_sistema.

    """

    if len(qr) < 6:

        return ""


    # Extrair número de referência dos primeiros 6 dígitos

    referencia = qr[:6]

    try:

        referencia_num = int(referencia.lstrip('0'))

    except ValueError:

        return ""


    # Verificar se os últimos 3 caracteres são 'MPI'

    if not qr[-3:] == "MPI":

        return ""


    # Buscar no banco

    try:

        cursor = conn.cursor()

        cursor.execute(

            "SELECT Desc_Proc_01, Tipo FROM mp_sistema WHERE codigo_01 = ?",

            (referencia_num,)

        )

        row = cursor.fetchone()

        cursor.close()


        # Só retorna se Tipo for exatamente "Emulsão"

        if row and row[1] == "Emulsão":

            return row[0]  # Retorna Desc_Proc_01

        else:

            return ""

    except Exception as e:

        print(f"Erro ao buscar emulsao: {e}")

        return ""


def validar_emulsao_qr(event):

    qr = entry_emulsao_qr.get().strip()

    emulsao = buscar_emulsao_por_qr(qr)

    entry_emulsao_qr_readonly.config(state="normal")

    entry_emulsao_qr_readonly.delete(0, "end")

    entry_emulsao_qr_readonly.insert(0, emulsao if emulsao else "Não encontrado")

    entry_emulsao_qr_readonly.config(state="readonly")


entry_emulsao_qr.bind("<FocusOut>", validar_emulsao_qr)


if db_online:

    preencher_parametros_esticagem()

    preencher_parametros_pos_esticagem()

    carregar_parametros()


def carregar_espessura_pos_esticagem(id_quadro):

    """

    Busca o valor mais recente de espessura_pos_esticagem (INT) para o id_quadro

    e preenche o Entry entry_espessura_1.

    """

    print("DEBUG: Função chamada com id_quadro =", id_quadro)  # 1. Verifica chamada


    if not id_quadro:

        print("DEBUG: id_quadro é None ou vazio")

        return


    try:

        cursor = conn.cursor()


        # Query com CAST para evitar problemas de tipo

        cursor.execute("""

            SELECT TOP 1 [espessura_pos_esticagem]

            FROM [pos_esticagem]

            WHERE [id_quadro] = CAST(? AS VARCHAR)

            ORDER BY [data] DESC

        """, (id_quadro,))


        row = cursor.fetchone()

        print("DEBUG: row retornado do banco:", row)  # 2. Verifica retorno do banco


        if row:

            pos_esticagem_valor = row[0]  # INT vindo do banco

        else:

            pos_esticagem_valor = ""

            print("DEBUG: Nenhum registro encontrado para esse id_quadro")


        # Atualiza Entry

        print("DEBUG: Valor a inserir no Entry:", pos_esticagem_valor)  # 3. Verifica valor que será inserido

        entry_espessura_1.config(state="normal")

        entry_espessura_1.delete(0, tk.END)

        entry_espessura_1.insert(0, str(pos_esticagem_valor))

        entry_espessura_1.config(state="readonly")


        cursor.close()

        print("DEBUG: Entry atualizado com sucesso")


    except Exception as e:

        print("DEBUG: Erro na query")

        messagebox.showerror(

            "Erro de Banco",

            f"Ocorreu um erro ao buscar espessura:\n{e}"

        )


def verificar_reutilizacao(id_quadro):

    """

    Verifica se o quadro pode ser reutilizado.

    - Se reutilizar = 'Não' na tabela revelacao_final, bloqueia o carregamento.

    - Caso contrário, permite.

    """

    try:

        with conn.cursor() as cursor:

            cursor.execute("""

                SELECT TOP 1 reutilizar

                FROM revelacao_final

                WHERE UPPER(id_quadro) = ?

                ORDER BY data DESC

            """, id_quadro.upper())

            row = cursor.fetchone()


        if row and row[0] and row[0].strip().lower() == "não":

            messagebox.showwarning("Aviso", "Esse quadro não pode ser reutilizado")

            limpar_id_quadro()

            return False  # bloqueia


        return True  # permitido


    except Exception as e:

        messagebox.showerror("Erro", f"Falha ao verificar reutilização: {e}")

        return False


def ao_perder_foco(event=None):

    id_quadro_val = entry_id_quadro.get().strip()

    if not id_quadro_val:

        return


    # --- Verifica primeiro se o quadro pode ser reutilizado ---

    if not verificar_reutilizacao(id_quadro_val):

        return  # sai sem carregar nada


    # --- Só depois valida e segue com os carregamentos ---

    validar_e_carregar_quadro()

    carregar_espessura_pos_esticagem(id_quadro_val)

    carregar_emulsao_por_quadro(id_quadro_val)

    carregar_revelacao_por_quadro(id_quadro_val)


# Bind correto usando apenas a função existente

entry_id_quadro.bind("<FocusOut>", ao_perder_foco)

entry_id_quadro.focus_set()


root.mainloop()


