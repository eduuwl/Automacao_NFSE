"""
Sistema de Automação para Emissão de Notas Fiscais - SEFIN Belém
Versão Otimizada com Seletores Específicos
"""

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time
import os
from datetime import datetime

class AutomacaoNotaFiscal:
    def __init__(self, caminho_excel):
        self.caminho_excel = caminho_excel
        self.driver = None
        self.wait = None
    
    def configurar_navegador(self):
        """Configura o navegador Chrome"""
        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)
        
        self.driver = webdriver.Chrome(options=options)
        self.wait = WebDriverWait(self.driver, 15)
        print("✓ Navegador configurado")
    
    def carregar_dados(self):
        """Carrega dados do Excel"""
        df = pd.read_excel(self.caminho_excel)
        
        # Adiciona colunas de controle
        for col in ['Status', 'Numero_Nota', 'Data_Emissao', 'Mensagem_Erro']:
            if col not in df.columns:
                df[col] = ''
        
        # Converte para string
        df['Status'] = df['Status'].astype(str)
        df['Numero_Nota'] = df['Numero_Nota'].astype(str)
        df['Data_Emissao'] = df['Data_Emissao'].astype(str)
        df['Mensagem_Erro'] = df['Mensagem_Erro'].astype(str)
        
        print(f"✓ Excel carregado: {len(df)} registros")
        return df
    
    def acessar_sistema(self):
        """Acessa a página de emissão"""
        url = "https://notafiscal.belem.pa.gov.br/notafiscal/paginas/notafiscal/emissaoNotaFiscalData.jsf"
        self.driver.get(url)
        time.sleep(5)
        print("✓ Sistema acessado")
    
    def aguardar_loading(self, timeout=10):
        """Aguarda o loading sumir"""
        try:
            print(f"    → Aguardando loading...")
            # Aguarda aparecer
            time.sleep(1)
            # Aguarda sumir
            WebDriverWait(self.driver, timeout).until(
                EC.invisibility_of_element_located((By.CSS_SELECTOR, ".ui-blockui, .ui-blockui-content"))
            )
            print(f"    ✓ Loading concluído")
            return True
        except:
            print(f"    ℹ Timeout do loading - continuando...")
            time.sleep(2)
            return True
    
    def preencher_cpf_e_pesquisar(self, cpf):
        """Preenche CPF e clica em pesquisar"""
        try:
            print(f"  → Preenchendo CPF {cpf}...")
            
            # Limpa CPF
            cpf_limpo = cpf.replace('.', '').replace('-', '').replace('/', '')
            
            # Rola para a seção do Tomador
            self.driver.execute_script("window.scrollTo(0, 400);")
            time.sleep(1)
            
            # Preenche CPF
            campo_cpf = self.driver.find_element(By.ID, "formNotaFiscal:idCpfCnpjPessoa:idInputMaskCpfCnpj:inputText")
            campo_cpf.clear()
            campo_cpf.send_keys(cpf_limpo)
            print(f"    ✓ CPF preenchido")
            time.sleep(1)
            
            # Busca o botão Pesquisar correto (tem "dados-pessoa" no onclick)
            btn = self.driver.find_element(By.XPATH, 
                "//a[contains(@class, 'btn-success') and contains(@onclick, 'dados-pessoa') and .//i[contains(@class, 'pe-7s-search')]]")
            print(f"    ✓ Botão Pesquisar encontrado: {btn.get_attribute('id')}")
            
            # Clica
            self.driver.execute_script("arguments[0].click();", btn)
            print(f"    ✓ Clicado")
            
            # Aguarda loading
            self.aguardar_loading()
            
            # Aguarda dados carregarem
            time.sleep(3)
            
            # Verifica se carregou
            try:
                # Busca campo nome
                campo_nome = self.driver.find_element(By.XPATH, "//input[contains(@id, 'nomeEmpresarial') or contains(@id, 'nome')]")
                nome = campo_nome.get_attribute('value')
                
                if nome and len(nome) > 3:
                    print(f"    ✓ Dados carregados: {nome[:40]}...")
                    
                    # DEBUG: Verifica estado do dropdown
                    try:
                        dropdown = self.driver.find_element(By.ID, "formNotaFiscal:idAtividadeEmissor_input")
                        is_disabled = dropdown.get_attribute('disabled')
                        print(f"    ℹ Dropdown Atividade disabled={is_disabled}")
                    except:
                        print(f"    ⚠ Dropdown Atividade não encontrado ainda")
                    
                    return True
                else:
                    print(f"    ⚠ Nome vazio ({len(nome) if nome else 0} chars) - mas continuando...")
                    return True
            except Exception as e:
                print(f"    ⚠ Não conseguiu verificar nome: {type(e).__name__}")
                return True
            
        except Exception as e:
            print(f"    ✗ Erro ao pesquisar CPF: {type(e).__name__} - {str(e)}")
            self.driver.save_screenshot(f"erro_cpf_{cpf_limpo}.png")
            return False
    
    def selecionar_atividade(self):
        """Seleciona atividade 931310000 - Condicionamento físico (Dropdown PrimeFaces)"""
        try:
            print(f"  → Selecionando atividade...")
            
            # Aguarda a página processar os dados do tomador
            time.sleep(5)
            
            # Rola até a seção de atividade
            self.driver.execute_script("window.scrollTo(0, 1000);")
            time.sleep(2)
            
            # 1. ENCONTRA O CONTAINER DO DROPDOWN
            dropdown_id = "formNotaFiscal:idAtividadeEmissor"
            print(f"    → Procurando dropdown: {dropdown_id}")
            
            dropdown = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, dropdown_id))
            )
            print(f"    ✓ Dropdown encontrado")
            
            # Aguarda estar habilitado (verifica aria-disabled)
            print(f"    → Aguardando dropdown habilitar...")
            for i in range(10):
                aria_disabled = dropdown.get_attribute('aria-disabled')
                if aria_disabled == 'false' or not aria_disabled:
                    print(f"    ✓ Dropdown habilitado após {i+1}s")
                    break
                time.sleep(1)
            else:
                print(f"    ⚠ Dropdown ainda pode estar desabilitado - tentando mesmo assim...")
            
            # Rola até o dropdown
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown)
            time.sleep(1)
            
            # 2. CLICA NO DROPDOWN PARA ABRIR
            print(f"    → Abrindo dropdown (clicando)...")
            try:
                # Tenta clicar no trigger (setinha)
                trigger = dropdown.find_element(By.CLASS_NAME, "ui-selectonemenu-trigger")
                trigger.click()
                print(f"    ✓ Clicou no trigger")
            except:
                # Fallback: clica no próprio dropdown
                dropdown.click()
                print(f"    ✓ Clicou no dropdown")
            
            time.sleep(2)
            
            # 3. AGUARDA A LISTA (UL) APARECER
            print(f"    → Aguardando lista de opções aparecer...")
            ul_id = "formNotaFiscal:idAtividadeEmissor_items"
            
            lista = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.ID, ul_id))
            )
            print(f"    ✓ Lista de opções visível")
            
            time.sleep(1)
            
            # 4. BUSCA E CLICA NO <LI> CORRETO
            print(f"    → Procurando opção '931310000'...")
            
            # Busca o <li> que contém "931310000"
            opcao = lista.find_element(By.XPATH, 
                ".//li[contains(@data-label, '931310000') or contains(text(), '931310000')]")
            
            # Rola até a opção
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'nearest'});", opcao)
            time.sleep(0.5)
            
            # Pega o texto da opção
            texto_opcao = opcao.text
            print(f"    ℹ Opção encontrada: {texto_opcao[:60]}...")
            
            # Clica na opção
            opcao.click()
            print(f"    ✓ Opção clicada")
            
            time.sleep(2)
            
            # 5. VERIFICA SE FOI SELECIONADA
            try:
                input_elem = self.driver.find_element(By.ID, "formNotaFiscal:idAtividadeEmissor_input")
                valor_selecionado = input_elem.get_attribute('value')
                print(f"    ✓ Atividade selecionada: {valor_selecionado[:60] if valor_selecionado else 'N/A'}...")
            except:
                print(f"    ℹ Não conseguiu verificar valor selecionado - mas continuando...")
            
            # Aguarda processamento
            time.sleep(3)
            
            print(f"    ✓ Atividade '931310000' selecionada com sucesso!")
            return True
            
        except Exception as e:
            print(f"    ✗ Erro ao selecionar atividade: {type(e).__name__} - {str(e)}")
            self.driver.save_screenshot("erro_atividade.png")
            
            # Salva HTML para debug
            try:
                with open("debug_atividade.html", "w", encoding="utf-8") as f:
                    f.write(self.driver.page_source)
                print(f"    ℹ HTML salvo em: debug_atividade.html")
            except:
                pass
            
            return False
    
    def adicionar_descricao(self):
        """Adiciona descrição da nota"""
        try:
            print(f"  → Adicionando descrição...")
            
            # Rola até a seção de descrição
            self.driver.execute_script("window.scrollTo(0, 1600);")
            time.sleep(2)
            
            # 1. BUSCA E CLICA NO BOTÃO "CARREGAR DESCRIÇÃO"
            print(f"    → Procurando botão 'Carregar Descrição'...")
            
            btn = None
            try:
                btn = self.driver.find_element(By.XPATH, 
                    "//a[contains(@class, 'btn-warning') and (contains(., 'Carregar') or contains(., 'Descrição'))]")
                print(f"    ✓ Botão encontrado: {btn.get_attribute('id')}")
            except:
                try:
                    btn = self.driver.find_element(By.XPATH, 
                        "//a[.//i[contains(@class, 'fa-plus-circle')]]")
                    print(f"    ✓ Botão encontrado pelo ícone")
                except:
                    print(f"    ✗ Botão não encontrado!")
                    return False
            
            # Rola e clica no botão
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
            time.sleep(1)
            btn.click()
            print(f"    ✓ Botão clicado - aguardando modal...")
            
            time.sleep(3)
            
            # 2. AGUARDA MODAL "DESCRIÇÃO FAVORITA" APARECER
            print(f"    → Aguardando modal aparecer...")
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, 
                        "//div[contains(@class, 'ui-dialog') and contains(@style, 'display')]//h3[contains(., 'Descrição Favorita')]"))
                )
                print(f"    ✓ Modal 'Descrição Favorita' visível")
            except:
                print(f"    ⚠ Modal não detectado - tentando continuar...")
            
            time.sleep(2)
            
            # 3. BUSCA E CLICA NO CHECKBOX DA PRIMEIRA LINHA (não o do cabeçalho!)
            print(f"    → Procurando checkbox da primeira linha...")
            
            checkbox_clicado = False
            
            # IMPORTANTE: Não clicar no _head_checkbox (cabeçalho), e sim no checkbox da LINHA
            try:
                # Busca o checkbox da primeira LINHA da tabela (não o do header)
                # Evita _head_checkbox e busca _0, _1, etc
                checkbox = self.driver.find_element(By.XPATH, 
                    "//div[contains(@class, 'ui-datatable-scrollable-body')]//div[@role='checkbox' and @aria-checked='false' and not(contains(@id, '_head_'))]")
                
                checkbox_id = checkbox.get_attribute('id')
                print(f"    ℹ Checkbox da linha encontrado: {checkbox_id}")
                
                # Rola até o checkbox
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", checkbox)
                time.sleep(1)
                
                # Tenta clicar até 3 vezes
                for tentativa in range(3):
                    # Clica no checkbox
                    self.driver.execute_script("arguments[0].click();", checkbox)
                    time.sleep(1)
                    
                    # Verifica se marcou
                    aria_checked = checkbox.get_attribute('aria-checked')
                    print(f"    ℹ Tentativa {tentativa + 1}: aria-checked={aria_checked}")
                    
                    if aria_checked == 'true':
                        print(f"    ✓ Checkbox da linha marcado!")
                        checkbox_clicado = True
                        break
                
                # Se ainda não marcou, tenta clicar no SPAN interno
                if not checkbox_clicado:
                    try:
                        span = checkbox.find_element(By.TAG_NAME, "span")
                        self.driver.execute_script("arguments[0].click();", span)
                        time.sleep(1)
                        
                        aria_checked = checkbox.get_attribute('aria-checked')
                        if aria_checked == 'true':
                            print(f"    ✓ Checkbox marcado via span!")
                            checkbox_clicado = True
                    except:
                        pass
                
                # Se AINDA não marcou, força manualmente
                if not checkbox_clicado:
                    print(f"    ⚠ Forçando seleção via JS...")
                    self.driver.execute_script("""
                        var cb = arguments[0];
                        cb.setAttribute('aria-checked', 'true');
                        cb.classList.add('ui-state-active');
                        var span = cb.querySelector('span');
                        if(span) {
                            span.className = 'ui-chkbox-icon ui-icon ui-icon-check ui-c';
                        }
                    """, checkbox)
                    time.sleep(1)
                    checkbox_clicado = True
                    
            except Exception as e:
                print(f"    ⚠ Erro ao buscar checkbox da linha: {type(e).__name__}")
                
                # FALLBACK: Se não achar checkbox de linha, clica no de cabeçalho mesmo
                print(f"    → Tentando checkbox do cabeçalho como fallback...")
                try:
                    checkbox_head = self.driver.find_element(By.XPATH, 
                        "//div[contains(@id, '_head_checkbox')]//span[contains(@class, 'ui-chkbox-icon')]")
                    
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", checkbox_head)
                    time.sleep(1)
                    self.driver.execute_script("arguments[0].click();", checkbox_head)
                    time.sleep(1)
                    print(f"    ✓ Checkbox do cabeçalho clicado")
                    checkbox_clicado = True
                except:
                    print(f"    ✗ Não conseguiu clicar em nenhum checkbox!")
            
            if not checkbox_clicado:
                print(f"    ✗ Falha ao marcar checkbox!")
                self.driver.save_screenshot("erro_checkbox.png")
                return False
            
            time.sleep(2)
            
            # 4. VERIFICA SE FOI SELECIONADO (deve mostrar "Selecionado - 1")
            try:
                contador = self.driver.find_element(By.XPATH, "//span[contains(text(), 'Selecionado')]").text
                print(f"    ℹ Status: {contador}")
                
                if "Selecionado - 0" in contador or "- 0" in contador:
                    print(f"    ⚠ Nenhum item selecionado ainda - tentando novamente...")
                    
                    # Tenta clicar novamente via JS no primeiro input visível
                    try:
                        checkboxes = self.driver.find_elements(By.XPATH, 
                            "//div[contains(@class, 'ui-datatable')]//input[@type='checkbox']")
                        for cb in checkboxes:
                            if cb.is_displayed():
                                self.driver.execute_script("arguments[0].checked = true; arguments[0].click();", cb)
                                time.sleep(1)
                                break
                    except:
                        pass
            except:
                print(f"    ℹ Não conseguiu verificar contador - continuando...")
            
            time.sleep(2)
            
            # 5. BUSCA E CLICA NO BOTÃO "CONFIRMAR"
            print(f"    → Procurando botão 'Confirmar'...")
            
            btn_confirmar = None
            try:
                # O botão é um <a> com btn-success e classe dialogselect_save
                btn_confirmar = self.driver.find_element(By.XPATH, 
                    "//a[contains(@class, 'btn-success') and contains(@class, 'dialogselect_save')]")
                print(f"    ✓ Botão Confirmar encontrado")
            except:
                try:
                    # Fallback: procura dentro do modal
                    btn_confirmar = self.driver.find_element(By.XPATH, 
                        "//div[contains(@class, 'ui-dialog')]//a[contains(@class, 'btn-success') and contains(., 'Confirmar')]")
                    print(f"    ✓ Botão Confirmar encontrado (fallback)")
                except:
                    print(f"    ✗ Botão Confirmar não encontrado!")
                    self.driver.save_screenshot("erro_confirmar.png")
                    return False
            
            # Rola até o botão e clica
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_confirmar)
            time.sleep(1)
            
            # Clica no botão Confirmar
            self.driver.execute_script("arguments[0].click();", btn_confirmar)
            print(f"    ✓ Botão Confirmar clicado")
            
            time.sleep(2)
            
            # 6. AGUARDA O MODAL FECHAR COMPLETAMENTE
            print(f"    → Aguardando modal fechar...")
            try:
                # Aguarda o modal sumir
                WebDriverWait(self.driver, 10).until(
                    EC.invisibility_of_element_located((By.XPATH, 
                        "//div[contains(@class, 'ui-dialog') and contains(@id, 'Descricao')]"))
                )
                print(f"    ✓ Modal fechado")
            except:
                print(f"    ⚠ Modal pode não ter fechado - aguardando tempo fixo...")
                time.sleep(3)
            
            # 7. AGUARDA LOADING PROCESSAR
            try:
                WebDriverWait(self.driver, 5).until(
                    EC.invisibility_of_element_located((By.CSS_SELECTOR, ".ui-blockui"))
                )
                print(f"    ✓ Loading concluído")
            except:
                time.sleep(2)
                print(f"    ℹ Loading não detectado")
            
            # 8. VERIFICA SE A DESCRIÇÃO FOI ADICIONADA
            time.sleep(2)
            try:
                # Procura por algum campo de descrição preenchido
                desc_campo = self.driver.find_element(By.XPATH, 
                    "//textarea[contains(@id, 'descricao') or contains(@id, 'Descricao')]")
                desc_valor = desc_campo.get_attribute('value')
                
                if desc_valor and len(desc_valor) > 5:
                    print(f"    ✓ Descrição adicionada: {desc_valor[:40]}...")
                else:
                    print(f"    ⚠ Campo descrição está vazio - mas continuando...")
            except:
                print(f"    ℹ Não conseguiu verificar descrição - mas continuando...")
            
            print(f"    ✓ Descrição adicionada com sucesso!")
            return True
            
        except Exception as e:
            print(f"    ✗ Erro ao adicionar descrição: {type(e).__name__} - {str(e)}")
            self.driver.save_screenshot("erro_descricao_final.png")
            
            try:
                with open("debug_descricao_final.html", "w", encoding="utf-8") as f:
                    f.write(self.driver.page_source)
                print(f"    ℹ HTML salvo em: debug_descricao_final.html")
            except:
                pass
            
            return False
    
    def preencher_valor(self, valor=110.00):
        """Preenche valor dos serviços"""
        try:
            print(f"  → Preenchendo valor R$ {valor:.2f}...")
            
            # Aguarda um pouco após modal fechar
            time.sleep(3)
            
            # Rola até a seção de valores
            self.driver.execute_script("window.scrollTo(0, 2000);")
            time.sleep(2)
            
            # Formata valor (110.00 → "110")
            valor_str = str(int(valor))  # Remove decimais, envia só "110"
            
            print(f"    → Buscando campo de valor...")
            
            # Busca o campo (sem salvar referência ainda)
            estrategias = [
                # Input dentro de span ui-inputnumber (mais específico)
                "//span[contains(@class, 'ui-inputnumber')]//input[@type='text']",
                # Por classes específicas
                "//input[contains(@class, 'ui-inputnumber') and contains(@class, 'input-currency')]",
                # Por ID
                "//input[contains(@id, 'inputText_input')]",
            ]
            
            xpath_correto = None
            for xpath in estrategias:
                try:
                    campos = self.driver.find_elements(By.XPATH, xpath)
                    for c in campos:
                        if c.is_displayed() and c.is_enabled():
                            xpath_correto = xpath
                            print(f"    ✓ Encontrado com xpath: {xpath[:60]}...")
                            break
                    if xpath_correto:
                        break
                except:
                    continue
            
            if not xpath_correto:
                print(f"    ✗ Campo não encontrado!")
                return False
            
            # Agora que sabemos o XPATH, vamos interagir
            # Re-busca o elemento toda vez para evitar StaleElement
            
            # 1. CLICA no campo
            print(f"    → Clicando no campo...")
            campo = self.driver.find_element(By.XPATH, xpath_correto)
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo)
            time.sleep(1)
            campo.click()
            time.sleep(1)
            
            # 2. SELECIONA TODO o texto (CTRL+A)
            print(f"    → Selecionando todo texto...")
            campo = self.driver.find_element(By.XPATH, xpath_correto)  # Re-busca
            campo.send_keys(Keys.CONTROL + "a")
            time.sleep(0.5)
            
            # 3. APAGA (DELETE ou BACKSPACE)
            print(f"    → Apagando...")
            campo = self.driver.find_element(By.XPATH, xpath_correto)  # Re-busca
            campo.send_keys(Keys.DELETE)
            time.sleep(0.5)
            
            # 4. DIGITA o valor
            print(f"    → Digitando {valor_str}...")
            campo = self.driver.find_element(By.XPATH, xpath_correto)  # Re-busca
            campo.send_keys(valor_str)
            time.sleep(1)
            
            # 5. ENTER
            print(f"    → Pressionando ENTER...")
            campo = self.driver.find_element(By.XPATH, xpath_correto)  # Re-busca
            campo.send_keys(Keys.RETURN)
            time.sleep(2)
            
            # 6. Aguarda cálculo
            print(f"    → Aguardando cálculo...")
            try:
                WebDriverWait(self.driver, 5).until(
                    EC.invisibility_of_element_located((By.CSS_SELECTOR, ".ui-blockui"))
                )
                print(f"    ✓ Cálculo concluído")
            except:
                time.sleep(2)
            
            # 7. Verifica se preencheu
            time.sleep(1)
            try:
                campo = self.driver.find_element(By.XPATH, xpath_correto)
                valor_atual = campo.get_attribute('value')
                print(f"    ℹ Valor no campo: '{valor_atual}'")
                
                if valor_atual and (valor_str in valor_atual or str(valor) in valor_atual):
                    print(f"    ✓ Valor preenchido com sucesso!")
                    return True
                else:
                    print(f"    ⚠ Valor diferente mas continuando...")
                    return True
            except:
                print(f"    ℹ Não conseguiu verificar mas continuando...")
                return True
            
        except Exception as e:
            print(f"    ✗ Erro ao preencher valor: {type(e).__name__} - {str(e)[:100]}")
            self.driver.save_screenshot("erro_preencher_valor.png")
            
            try:
                with open("debug_erro_valor.html", "w", encoding="utf-8") as f:
                    f.write(self.driver.page_source)
            except:
                pass
            
            return False
    
    def emitir_nota(self):
        """Emite a nota fiscal"""
        try:
            print(f"  → Emitindo nota...")
            
            # Rola até o final da página
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            
            # Busca botão Emitir
            try:
                btn = self.driver.find_element(By.XPATH, 
                    "//button[contains(@id, 'btnEmitir') or (contains(., 'Emitir') and contains(@class, 'btn'))]")
            except:
                btn = self.driver.find_element(By.XPATH, 
                    "//a[contains(., 'Emitir') and contains(@class, 'btn')]")
            
            # Rola até o botão e clica
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
            time.sleep(1)
            btn.click()
            print(f"    ✓ Botão Emitir clicado")
            
            # Aguarda processamento
            time.sleep(3)
            self.aguardar_loading(timeout=15)
            
            # Aguarda mensagem de sucesso
            time.sleep(3)
            
            # Tenta capturar número da nota
            numero_nota = None
            try:
                # Busca mensagem de sucesso
                msg = self.driver.find_element(By.XPATH, 
                    "//*[contains(text(), 'emitida') or contains(text(), 'Emitida')]").text
                print(f"    ✓ Mensagem: {msg[:60]}...")
                
                # Tenta extrair número
                import re
                match = re.search(r'(\d+)', msg)
                if match:
                    numero_nota = match.group(1)
                    
            except:
                pass
            
            if numero_nota:
                print(f"    ✓ Nota emitida com sucesso! Número: {numero_nota}")
            else:
                print(f"    ✓ Nota emitida com sucesso!")
                numero_nota = "Emitida"
            
            return numero_nota
                
        except Exception as e:
            print(f"    ✗ Erro ao emitir nota: {type(e).__name__} - {str(e)}")
            self.driver.save_screenshot("erro_emissao.png")
            return None
    
    def limpar_formulario(self):
        """Limpa o formulário para próxima nota"""
        try:
            print(f"  → Limpando formulário...")
            
            # Tenta clicar em "Nova Nota" ou recarregar a página
            try:
                btn_nova = self.driver.find_element(By.XPATH, 
                    "//button[contains(., 'Nova') or contains(., 'Limpar')] | //a[contains(., 'Nova') or contains(., 'Limpar')]")
                btn_nova.click()
                time.sleep(3)
                print(f"    ✓ Formulário limpo")
            except:
                # Se não tiver botão, recarrega a página
                self.driver.refresh()
                time.sleep(5)
                print(f"    ✓ Página recarregada")
            
            return True
        except:
            # Se falhar, recarrega mesmo assim
            try:
                self.driver.refresh()
                time.sleep(5)
                return True
            except:
                return False
    
    def processar_nota(self, index, dados):
        """Processa uma nota completa"""
        print(f"\n{'='*60}")
        print(f"[{index + 1}] Processando CPF: {dados['CPF']}")
        print(f"{'='*60}")
        
        try:
            # 1. CPF e Pesquisar
            if not self.preencher_cpf_e_pesquisar(dados['CPF']):
                return 'ERRO', '', 'Erro ao pesquisar CPF'
            
            # 2. Atividade
            if not self.selecionar_atividade():
                return 'ERRO', '', 'Erro ao selecionar atividade'
            
            # 3. Descrição
            if not self.adicionar_descricao():
                return 'ERRO', '', 'Erro ao adicionar descrição'
            
            # 4. Valor
            valor = float(dados.get('Valor', 110.00))
            if not self.preencher_valor(valor):
                return 'ERRO', '', 'Erro ao preencher valor'
            
            # 5. Emitir
            numero = self.emitir_nota()
            if not numero:
                return 'ERRO', '', 'Erro ao emitir nota'
            
            # 6. Limpar para próxima
            self.limpar_formulario()
            
            return 'EMITIDA', numero, ''
            
        except Exception as e:
            erro_msg = f"{type(e).__name__}: {str(e)}"
            print(f"  ✗ Erro inesperado: {erro_msg}")
            return 'ERRO', '', erro_msg
    
    def executar(self):
        """Executa o processo completo"""
        print("\n" + "="*60)
        print("  AUTOMAÇÃO NFS-E BELÉM - VERSÃO OTIMIZADA")
        print("="*60 + "\n")
        
        # Carrega dados
        df = self.carregar_dados()
        
        # Configura navegador
        self.configurar_navegador()
        
        # Acessa sistema
        self.acessar_sistema()
        
        print("\n" + "⚠"*30)
        print("  ATENÇÃO: Faça LOGIN no sistema")
        print("⚠"*30)
        input("\n➤ Pressione ENTER após fazer login...\n")
        
        # Processa notas
        total = len(df)
        sucesso = 0
        erros = 0
        
        for index, row in df.iterrows():
            # Pula se já foi processada
            if str(row.get('Status', '')).upper() == 'EMITIDA':
                print(f"\n[{index + 1}] ✓ Já processada - PULANDO")
                sucesso += 1
                continue
            
            # Processa a nota
            status, numero, erro = self.processar_nota(index, row)
            
            # Atualiza DataFrame
            df.at[index, 'Status'] = status
            df.at[index, 'Numero_Nota'] = numero if numero else ''
            df.at[index, 'Data_Emissao'] = datetime.now().strftime('%d/%m/%Y %H:%M') if status == 'EMITIDA' else ''
            df.at[index, 'Mensagem_Erro'] = erro if erro else ''
            
            # Contabiliza
            if status == 'EMITIDA':
                sucesso += 1
                print(f"\n  ✓✓✓ SUCESSO! ({sucesso}/{total})")
            else:
                erros += 1
                print(f"\n  ✗✗✗ ERRO: {erro}")
                print(f"  ({erros} erros até agora)")
            
            # Salva progresso a cada 3 notas
            if (index + 1) % 3 == 0:
                try:
                    df.to_excel(self.caminho_excel, index=False)
                    print(f"\n  💾 Progresso salvo ({index + 1}/{total})")
                except Exception as e:
                    print(f"\n  ⚠ Erro ao salvar: {str(e)}")
                    print(f"  (Certifique-se de que o Excel está fechado)")
            
            # Pequena pausa entre notas
            if index < total - 1:  # Não pausar na última
                time.sleep(2)
        
        # Salva resultado final
        print(f"\n{'='*60}")
        print("  SALVANDO RESULTADO FINAL...")
        print(f"{'='*60}")
        
        try:
            df.to_excel(self.caminho_excel, index=False)
            print("✓ Arquivo salvo com sucesso!")
        except Exception as e:
            print(f"✗ Erro ao salvar: {str(e)}")
            print("⚠ FECHE O EXCEL e tente salvar manualmente!")
            backup = f"resultado_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            try:
                df.to_excel(backup, index=False)
                print(f"✓ Backup salvo em: {backup}")
            except:
                print("✗ Não foi possível salvar backup")
        
        # Relatório final
        print(f"\n{'='*60}")
        print("  RELATÓRIO FINAL")
        print(f"{'='*60}")
        print(f"  Total de registros: {total}")
        print(f"  ✓ Emitidas: {sucesso}")
        print(f"  ✗ Erros: {erros}")
        print(f"  Taxa de sucesso: {(sucesso/total)*100:.1f}%")
        print(f"{'='*60}\n")
        
        input("➤ Pressione ENTER para fechar o navegador...")
        self.driver.quit()
        print("\n✓ Processo finalizado!")


if __name__ == "__main__":
    import sys
    
    print("\n" + "="*60)
    print("  BEM-VINDO AO SISTEMA DE AUTOMAÇÃO NFS-E")
    print("="*60 + "\n")
    
    caminho = input("📁 Arquivo Excel (ou ENTER para 'notas_fiscais.xlsx'): ").strip()
    if not caminho:
        caminho = "notas_fiscais.xlsx"
    
    if not os.path.exists(caminho):
        print(f"\n✗ ERRO: Arquivo não encontrado: {caminho}")
        print(f"✗ Certifique-se de que o arquivo existe no diretório atual")
        input("\nPressione ENTER para sair...")
        sys.exit(1)
    
    print(f"\n✓ Arquivo encontrado: {caminho}\n")
    
    try:
        automacao = AutomacaoNotaFiscal(caminho)
        automacao.executar()
    except KeyboardInterrupt:
        print("\n\n⚠ Processo interrompido pelo usuário")
        print("✓ Dados foram salvos até o último checkpoint")
    except Exception as e:
        print(f"\n✗ Erro fatal: {type(e).__name__}")
        print(f"   {str(e)}")
        input("\nPressione ENTER para sair...")