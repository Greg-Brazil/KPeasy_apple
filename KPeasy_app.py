import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import openpyxl
import traceback 
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException

MAX_TENTATIVAS_POR_LINHA = 3
DRY_RUN = False
VERIFY_FIRST_ROW = True
options = Options()
options.page_load_strategy = "eager"

def verify_dropdown(navegador, heading_text, option_text, timeout=0.5):
    hid = WebDriverWait(navegador, 2).until(
        EC.presence_of_element_located((By.XPATH, f"//*[contains(normalize-space(),'{heading_text}')]/ancestor::*[starts-with(@id,'QuestionId_')][1]"))
    ).get_attribute("id")
    opener = WebDriverWait(navegador, 2).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, f"[aria-haspopup='listbox'][aria-labelledby*='{hid}']"))
    )
    WebDriverWait(navegador, timeout).until(lambda d: option_text.lower() in (opener.text or "").lower())

def select_dropdown_fast(navegador, heading_text: str, option_text: str):
    """
    Abre o dropdown ancorado no heading e seleciona a opção por JS (com fallbacks robustos).
    heading_text: texto visível do título da pergunta (ex.: 'Informe CDRC')
    option_text: texto exato da opção (ex.: 'CDRC-SAPUCAIA', 'RJ')
    """
    # 1) container da pergunta ⇒ id 'QuestionId_...'
    heading = WebDriverWait(navegador, 6).until(
        EC.presence_of_element_located((
            By.XPATH,
            f"//*[contains(normalize-space(),'{heading_text}')]/ancestor::*[starts-with(@id,'QuestionId_')][1]"
        ))
    )
    hid = heading.get_attribute("id")

    # 2) opener do combobox associado a esse heading
    opener = WebDriverWait(navegador, 6).until(
        EC.element_to_be_clickable((
            By.CSS_SELECTOR, f"[aria-haspopup='listbox'][aria-labelledby*='{hid}']"
        ))
    )

    # Se já parece selecionado, não gasta tempo
    try:
        if option_text.strip().lower() in (opener.text or "").strip().lower():
            return
    except Exception:
        pass

    try:
        navegador.execute_script("arguments[0].scrollIntoView({block:'center'});", opener)
    except Exception:
        pass
    opener.click()

    # 3) Fast-path por JS: acha a opção visual, rola e clica
    js = r"""
    const needle = (arguments[0]||'').trim().toLowerCase();
    const spans = Array.from(document.querySelectorAll("div[role='option'] span"));
    // 3a) match exato primeiro
    for (const s of spans) {
      const txt = (s.textContent||'').trim().toLowerCase();
      if (txt === needle) {
        s.scrollIntoView({block:'center'});
        s.closest("div[role='option']").click();
        return true;
      }
    }
    // 3b) se não achou, aceita 'contains' como fallback
    for (const s of spans) {
      const txt = (s.textContent||'').trim().toLowerCase();
      if (txt.includes(needle)) {
        s.scrollIntoView({block:'center'});
        s.closest("div[role='option']").click();
        return true;
      }
    }
    return false;
    """
    try:
        ok = navegador.execute_script(js, option_text)
    except Exception:
        ok = False

    if ok:
        return

    # 4) Fallback Selenium: espera "clicável" e clica
    try:
        opt = WebDriverWait(navegador, 4).until(
            EC.element_to_be_clickable((
                By.XPATH,
                f"//div[@role='option'][.//span[normalize-space()='{option_text}']]"
            ))
        )
        try:
            navegador.execute_script("arguments[0].scrollIntoView({block:'center'});", opt)
        except Exception:
            pass
        opt.click()
        return
    except TimeoutException:
        # 5) Último recurso: reabrir rapidamente e tentar de novo
        try:
            opener.click()
        except Exception:
            pass
        opt = WebDriverWait(navegador, 4).until(
            EC.element_to_be_clickable((
                By.XPATH,
                f"//div[@role='option'][.//span[normalize-space()='{option_text}']]"
            ))
        )
        opt.click()
        return


# Função para iniciar o preenchimento do formulário
def preencher_formulario():
    # Pega os valores inseridos pelo usuário
    nome_usuario = nome_entry.get()
    planilha_caminho = planilha_entry.get()

    if not nome_usuario or not planilha_caminho:
        messagebox.showerror("Erro", "Por favor, preencha todos os campos!")
        return
    
    try:
        # Lê a planilha Excel
        planilha = pd.read_excel(planilha_caminho)

        # Inicializa o navegador
        navegador = webdriver.Chrome(options=options)

        # Acesse o formulário (novo link atualizado)
        url_formulario = "https://forms.cloud.microsoft/Pages/ResponsePage.aspx?id=ND7BLz_wi0mYKny0RuJbxusFv2jY35FPpnPGVFvj0_FUQVVSMkwzNDQxRUtIVFVPT00wM0ZLN0RCVC4u"

        try:
            linhas_ok = []      # [(linha_excel, veterinario)]
            linhas_erro = []    # [(linha_excel, veterinario, msg_erro)]
            contagem_por_tipo = {}
            
            # Loop para preencher o formulário para cada linha da planilha
            for index, row in planilha.iterrows():
                linha_excel = index + 2
                vet_nome = str(row.get("Veterinario", "") or "").strip() or "—"

                for tentativa in range(1, MAX_TENTATIVAS_POR_LINHA + 1):
                    try:
                        print(f"Preenchendo o formulário com os dados da linha {index + 1}...")

                        # Abrir o formulário novamente para cada linha
                        navegador.get(url_formulario)
        
                        # --- CDRC (dropdown) ---
                        cdrc_val = "CDRC-SAPUCAIA"
                        print(f"Selecionando CDRC: {cdrc_val}")
                        select_dropdown_fast(navegador, "Informe CDRC", cdrc_val)
                        print("CDRC selecionado.")
        
                        # --- ESTADO (dropdown) ---
                        estado_val = "RJ"
                        print(f"Selecionando Estado: {estado_val}")
                        select_dropdown_fast(navegador, "Informe Estado", estado_val)
                        print("Estado selecionado.")

                        if VERIFY_FIRST_ROW and index == 0:
                            verify_dropdown(navegador, "Informe CDRC", cdrc_val)
                            verify_dropdown(navegador, "Informe Estado", estado_val)

                        # --- CONSULTOR TÉCNICO (nome) ---
                        print(f"Preenchendo nome do consultor: {nome_usuario}")
                        
                        xpath_nome = (
                            "("
                            "//div[@id='question-list']//div[.//*[contains(normalize-space(),'Consultor Técnico') and contains(normalize-space(),'nome')]]//input"
                            " | "
                            "//div[@id='question-list']//div[.//*[contains(normalize-space(),'Consultor Técnico') and contains(normalize-space(),'nome')]]//textarea"
                            ")"
                        )
                        
                        caixa_nome = WebDriverWait(navegador, 10).until(
                            EC.presence_of_element_located((By.XPATH, xpath_nome))
                        )
                        try:
                            navegador.execute_script("arguments[0].scrollIntoView({block:'center'});", caixa_nome)
                        except Exception:
                            pass
                        try:
                            caixa_nome.clear()
                        except Exception:
                            pass
                        caixa_nome.send_keys(nome_usuario)

                        print("Nome do consultor preenchido.")
        
        
                       # Ler o tipo desta linha da planilha
                        valor_tipo = row["Tipo Visita"]
                        print(f"Selecionando tipo de visita: {valor_tipo}")
                        
                        # Seleção robusta por texto → clicar o ancestral clicável do label
                        xpath_tipo_clickable = (
                            f"//div[@id='question-list']"
                            f"//*[normalize-space()='{valor_tipo}']"
                            f"/ancestor::*[self::label or @role='radio' or @role='option' or self::button][1]"
                        )
                        opcao_tipo = WebDriverWait(navegador, 10).until(
                            EC.element_to_be_clickable((By.XPATH, xpath_tipo_clickable))
                        )
                        try:
                            navegador.execute_script("arguments[0].scrollIntoView({block:'center'});", opcao_tipo)
                        except Exception:
                            pass
                        opcao_tipo.click()
                        print(f"Opção '{valor_tipo}' foi selecionada.")

                        
                        print(f"Opção '{valor_tipo}' foi selecionada.")
        
                        if valor_tipo == "Visita Vet":
                            # Preencher os outros campos do formulário
                            # --- VISITA VET: Nome do estabelecimento ---
                            clinica = str(row.get("Clinica Veterinaria", "")).strip()
                            if not clinica:
                                raise ValueError(f"Linha {index+2}: coluna 'Clinica Veterinaria' está vazia para Visita Vet.")
                            
                            print(f"Visita Vet → Nome do estabelecimento: {clinica}")
                            
                            xpath_vet_estab = (
                                "("
                                "//div[@id='question-list']//div[.//*[normalize-space()='Nome do estabelecimento:' or contains(normalize-space(),'Nome do estabelecimento')]]//input"
                                " | "
                                "//div[@id='question-list']//div[.//*[normalize-space()='Nome do estabelecimento:' or contains(normalize-space(),'Nome do estabelecimento')]]//textarea"
                                ")"
                            )
                            
                            campo_vet_estab = WebDriverWait(navegador, 10).until(
                                EC.presence_of_element_located((By.XPATH, xpath_vet_estab))
                            )
                            try:
                                navegador.execute_script("arguments[0].scrollIntoView({block:'center'});", campo_vet_estab)
                            except Exception:
                                pass
                            try:
                                campo_vet_estab.clear()
                            except Exception:
                                pass
                            campo_vet_estab.send_keys(clinica)

                            print("Nome do estabelecimento (Visita Vet) preenchido.")
        
            
                            # --- VISITA VET: Nome do veterinário ---
                            nome_vet = str(row.get("Veterinario", "")).strip()
                            if not nome_vet:
                                raise ValueError(f"Linha {index+2}: coluna 'Veterinario' está vazia para Visita Vet.")
                            
                            print(f"Visita Vet → Nome do veterinário: {nome_vet}")
                            
                            xpath_vet_nome = (
                                "("
                                "//div[@id='question-list']//div[.//*[contains(normalize-space(),'Nome do veterinário')]]//input"
                                " | "
                                "//div[@id='question-list']//div[.//*[contains(normalize-space(),'Nome do veterinário')]]//textarea"
                                ")"
                            )
                            
                            campo_vet_nome = WebDriverWait(navegador, 10).until(
                                EC.presence_of_element_located((By.XPATH, xpath_vet_nome))
                            )
                            try:
                                navegador.execute_script("arguments[0].scrollIntoView({block:'center'});", campo_vet_nome)
                            except Exception:
                                pass
                            try:
                                campo_vet_nome.clear()
                            except Exception:
                                pass
                            campo_vet_nome.send_keys(nome_vet)
                            print("Nome do veterinário (Visita Vet) preenchido.")
        
            
                            # --- VISITA VET: Tema abordado ---
                            tema_vet = str(row.get("Assunto", "")).strip()
                            if not tema_vet:
                                raise ValueError(f"Linha {index+2}: coluna 'Assunto' está vazia para Visita Vet.")
                            
                            print(f"Visita Vet → Tema abordado: {tema_vet}")
                            
                            xpath_vet_tema = (
                                "("
                                "//div[@id='question-list']//div[.//*[contains(normalize-space(),'Tema abordado')]]//input"
                                " | "
                                "//div[@id='question-list']//div[.//*[contains(normalize-space(),'Tema abordado')]]//textarea"
                                ")"
                            )
                            
                            campo_vet_tema = WebDriverWait(navegador, 10).until(
                                EC.presence_of_element_located((By.XPATH, xpath_vet_tema))
                            )
                            try:
                                navegador.execute_script("arguments[0].scrollIntoView({block:'center'});", campo_vet_tema)
                            except Exception:
                                pass
                            try:
                                campo_vet_tema.clear()
                            except Exception:
                                pass
                            campo_vet_tema.send_keys(tema_vet)
                            print("Tema abordado (Visita Vet) preenchido.")

                        elif valor_tipo == "Mini Meeting":
                            # --- MINI MEETING: selecionar "Modelo" (Online/Presencial) a partir da planilha ---
                            modelo = str(row.get("Modelo", "")).strip()
                            if modelo not in ("Online", "Presencial"):
                                raise ValueError(f"Linha {index+2}: coluna 'Modelo' deve ser 'Online' ou 'Presencial' para Mini Meeting.")
                            
                            print(f"Mini Meeting → Modelo: {modelo}")
                            
                            # Clicar o ancestral clicável do texto (label/div role=radio/option/button) dentro do formulário
                            xpath_modelo_clickable = (
                                f"//div[@id='question-list']"
                                f"//*[normalize-space()='{modelo}']"
                                f"/ancestor::*[self::label or @role='radio' or @role='option' or self::button][1]"
                            )
                            
                            opcao = WebDriverWait(navegador, 10).until(
                                EC.element_to_be_clickable((By.XPATH, xpath_modelo_clickable))
                            )
                            try:
                                navegador.execute_script("arguments[0].scrollIntoView({block:'center'});", opcao)
                            except Exception:
                                pass
                            opcao.click()
                            print(f"Modelo '{modelo}' selecionado.")
        
                            # --- MINI MEETING: Nome do estabelecimento ---
                            estab = str(row.get("Estabelecimento", "")).strip()
                            if not estab:
                                raise ValueError(f"Linha {index+2}: coluna 'Estabelecimento' está vazia para Mini Meeting.")
                            
                            print(f"Mini Meeting → Estabelecimento: {estab}")
                            
                            xpath_estab = (
                                "("
                                "//div[@id='question-list']//div[.//*[contains(normalize-space(),'Nome do estabelecimento')]]//input"
                                " | "
                                "//div[@id='question-list']//div[.//*[contains(normalize-space(),'Nome do estabelecimento')]]//textarea"
                                ")"
                            )
                            
                            campo_estab = WebDriverWait(navegador, 10).until(
                                EC.presence_of_element_located((By.XPATH, xpath_estab))
                            )
                            try:
                                navegador.execute_script("arguments[0].scrollIntoView({block:'center'});", campo_estab)
                            except Exception:
                                pass
                            try:
                                campo_estab.clear()
                            except Exception:
                                pass
                            campo_estab.send_keys(estab)

                            print("Nome do estabelecimento preenchido.")
        
                            # --- MINI MEETING: Número de Participantes ---
                            raw = row.get("N participantes", "")
                            s = str(raw).strip().replace(",", ".")
                            if s == "" or s.lower() == "nan":
                                raise ValueError(f"Linha {index+2}: coluna 'N participantes' está vazia para Mini Meeting.")
                            try:
                                num = int(float(s))  # 4.0 -> 4 ; "4" -> 4 ; "4,0" -> 4
                            except Exception:
                                raise ValueError(f"Linha {index+2}: valor inválido em 'N participantes': {raw!r}")
                            num_part = str(num) 
                            print(f"Mini Meeting → Nº de participantes: {num_part}")
                            
                            xpath_participantes = (
                                "("
                                "//div[@id='question-list']//div[.//*[contains(normalize-space(),'Número de Participantes')]]//input"
                                " | "
                                "//div[@id='question-list']//div[.//*[contains(normalize-space(),'Número de Participantes')]]//textarea"
                                ")"
                            )
                            
                            campo_part = WebDriverWait(navegador, 10).until(
                                EC.presence_of_element_located((By.XPATH, xpath_participantes))
                            )
                            try:
                                navegador.execute_script("arguments[0].scrollIntoView({block:'center'});", campo_part)
                            except Exception:
                                pass
                            try:
                                campo_part.clear()
                            except Exception:
                                pass
                            campo_part.send_keys(num_part)
                            print("Número de Participantes preenchido.")

                            # --- MINI MEETING: Tema ---
                            tema = str(row.get("Tema", "")).strip()
                            if not tema:
                                raise ValueError(f"Linha {index+2}: coluna 'Tema' está vazia para Mini Meeting.")
                            
                            print(f"Mini Meeting → Tema: {tema}")
                            
                            xpath_tema = (
                                "("
                                "//div[@id='question-list']//div[.//*[contains(normalize-space(),'Tema')]]//input"
                                " | "
                                "//div[@id='question-list']//div[.//*[contains(normalize-space(),'Tema')]]//textarea"
                                ")"
                            )
                            
                            campo_tema = WebDriverWait(navegador, 10).until(
                                EC.presence_of_element_located((By.XPATH, xpath_tema))
                            )
                            try:
                                navegador.execute_script("arguments[0].scrollIntoView({block:'center'});", campo_tema)
                            except Exception:
                                pass
                            try:
                                campo_tema.clear()
                            except Exception:
                                pass
                            campo_tema.send_keys(tema)
                            print("Tema preenchido.")


                        elif valor_tipo == "Treinamento CDRC":
                            # --- CDRC: Número de Participantes (inteiro apenas) ---
                            raw = row.get("N CDRC", "")
                            s = str(raw).strip().replace(",", ".")
                            if s == "" or s.lower() == "nan":
                                raise ValueError(f"Linha {linha_excel}: coluna 'N CDRC' está vazia para Treinamento CDRC.")
                            try:
                                num = int(float(s))  # aceita 4, 4.0, "4,0" -> 4
                            except Exception:
                                raise ValueError(f"Linha {linha_excel}: valor inválido em 'N CDRC': {raw!r}")
                            num_str = str(num)
                            print(f"Treinamento CDRC → Nº de participantes: {num_str}")
                        
                            xpath_nc = (
                                "("
                                "//div[@id='question-list']//div[.//*[contains(normalize-space(),'Número de Participantes')]]//input"
                                " | "
                                "//div[@id='question-list']//div[.//*[contains(normalize-space(),'Número de Participantes')]]//textarea"
                                ")"
                            )
                            campo_nc = WebDriverWait(navegador, 10).until(
                                EC.presence_of_element_located((By.XPATH, xpath_nc))
                            )
                            try:
                                navegador.execute_script("arguments[0].scrollIntoView({block:'center'});", campo_nc)
                            except Exception:
                                pass
                            try:
                                campo_nc.clear()
                            except Exception:
                                pass
                            campo_nc.send_keys(num_str)
                            print("Número de Participantes (CDRC) preenchido.")
                        
                            # --- CDRC: Tema abordado (texto livre) ---
                            tema_cdrc = str(row.get("Tema CDRC", "")).strip()
                            if not tema_cdrc:
                                raise ValueError(f"Linha {linha_excel}: coluna 'Tema CDRC' está vazia para Treinamento CDRC.")
                            print(f"Treinamento CDRC → Tema: {tema_cdrc}")
                        
                            xpath_tema_cdrc = (
                                "("
                                "//div[@id='question-list']//div[.//*[contains(normalize-space(),'Tema abordado')]]//input"
                                " | "
                                "//div[@id='question-list']//div[.//*[contains(normalize-space(),'Tema abordado')]]//textarea"
                                ")"
                            )
                            campo_tema_cdrc = WebDriverWait(navegador, 10).until(
                                EC.presence_of_element_located((By.XPATH, xpath_tema_cdrc))
                            )
                            try:
                                navegador.execute_script("arguments[0].scrollIntoView({block:'center'});", campo_tema_cdrc)
                            except Exception:
                                pass
                            try:
                                campo_tema_cdrc.clear()
                            except Exception:
                                pass
                            campo_tema_cdrc.send_keys(tema_cdrc)
                            print("Tema (CDRC) preenchido.")



                        # Se chegou aqui sem exception, conta como sucesso e sai do retry
                        contagem_por_tipo[valor_tipo] = contagem_por_tipo.get(valor_tipo, 0) + 1
                        linhas_ok.append((linha_excel, vet_nome))
                        break

                    except Exception as e:
                        msg_curta = f"{type(e).__name__}: {e}"
                        print(f"[ERRO] Linha {linha_excel} (tentativa {tentativa}): {msg_curta}")
                        if tentativa >= MAX_TENTATIVAS_POR_LINHA:
                            linhas_erro.append((linha_excel, vet_nome, msg_curta))
                        # recupera o formulário para próxima tentativa / próxima linha
                        try:
                            navegador.get(url_formulario)
                            WebDriverWait(navegador, 15).until(
                                EC.presence_of_element_located((By.XPATH, "//*[contains(normalize-space(),'Informe CDRC')]"))
                            )
                        except Exception:
                            pass  # não atrapalhar o próximo ciclo
                        # continua o for tentativa (se ainda houver outra)
                        continue

                else:
                    print("Não tem essa opção ainda")
                    continue
                    
                # --- Enviar (com confirmação, sem XPath absoluto) ---
                
                if not DRY_RUN:
                    print("Aguardando o botão de envio...")
                    botao_enviar = WebDriverWait(navegador, 15).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "[data-automation-id='submitButton']"))
                    )
                    try:
                        navegador.execute_script("arguments[0].scrollIntoView({block:'center'});", botao_enviar)
                    except Exception:
                        pass
                    try:
                        botao_enviar.click()
                    except Exception:
                        # fallback JS se o click normal não disparar
                        navegador.execute_script("arguments[0].click();", botao_enviar)

                    WebDriverWait(navegador, 15).until(EC.staleness_of(botao_enviar))
                
                    print("Aguardando confirmação de envio...")
                    WebDriverWait(navegador, 20).until(
                        EC.presence_of_element_located((
                            By.XPATH,
                            "//*[contains(.,'Obrigado') or contains(.,'Thanks') or contains(.,'Your response') or contains(.,'Resposta')]"
                        ))
                    )
                    print("Envio confirmado.")
                else:
                    print("[DRY-RUN] Simulando envio (não clicou em Enviar).")

            # --- Resumo final ---
            total_ok = len(linhas_ok)
            total_err = len(linhas_erro)
            
            # Contagem enviada por tipo (só o que realmente passou no try)
            ordem = ["Visita Vet", "Mini Meeting", "Treinamento CDRC"]
            enviados_por_tipo = " | ".join(
                f"{t}: {contagem_por_tipo.get(t, 0)}" for t in ordem if contagem_por_tipo.get(t, 0) > 0
            ) or "Nenhum"
            
            # (Opcional, mas útil) Contagem esperada na planilha
            vc = planilha["Tipo Visita"].value_counts(dropna=False)
            esperados_por_tipo = " | ".join(f"{t}: {int(vc.get(t, 0))}" for t in ordem if vc.get(t, 0) > 0) or "Nenhum"
            
            if total_err == 0:
                messagebox.showinfo(
                    "Resumo",
                    f"Todas as {total_ok} linhas foram preenchidas com sucesso.\n"
                    f"Enviados por tipo: {enviados_por_tipo}\n"
                    f"Esperados na planilha: {esperados_por_tipo}"
                )
            else:
                linhas_txt = ", ".join(str(l) for (l, _, _) in linhas_erro)
                vets_txt   = ", ".join(v or "—" for (_, v, _) in linhas_erro)
                messagebox.showwarning(
                    "Resumo",
                    f"As linhas {linhas_txt} com os veterinários {vets_txt} não foram enviadas.\n"
                    f"Sucesso: {total_ok} | Erros: {total_err}\n"
                    f"Enviados por tipo: {enviados_por_tipo}\n"
                    f"Esperados na planilha: {esperados_por_tipo}"
                )


        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
        finally:
            navegador.quit()
    
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao carregar a planilha: {e}")

# Função para abrir o dialogo e escolher a planilha
def selecionar_planilha():
    caminho = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    if caminho:
        planilha_entry.delete(0, tk.END)
        planilha_entry.insert(0, caminho)

# Interface gráfica
root = tk.Tk()
root.title("KPEasy 2: o inimigo agora é outro")

# Nome
nome_label = tk.Label(root, text="Nome:")
nome_label.pack(pady=5)
nome_entry = tk.Entry(root, width=50)
nome_entry.pack(pady=5)

# Planilha
planilha_label = tk.Label(root, text="Selecione a Planilha (Excel):")
planilha_label.pack(pady=5)
planilha_entry = tk.Entry(root, width=50)
planilha_entry.pack(pady=5)
planilha_button = tk.Button(root, text="Selecionar Planilha", command=selecionar_planilha)
planilha_button.pack(pady=5)

# Botão de Iniciar
iniciar_button = tk.Button(root, text="Iniciar Preenchimento", command=preencher_formulario)
iniciar_button.pack(pady=20)

# Rodando a interface
root.mainloop()
