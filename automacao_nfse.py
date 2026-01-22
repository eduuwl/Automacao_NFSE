import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import os
from datetime import datetime

class AutomacaoNotaFiscal:
    def __init__(self, caminho_excel, pasta_downloads="notas_fiscais"):
        """
        Inicializa o sistema de automação
        
        Args:
            caminho_excel: Caminho para o arquivo Excel com os dados
            pasta_downloads: Pasta onde serão salvos os PDFs das notas
        """
        self.caminho_excel = caminho_excel
        self.pasta_downloads = pasta_downloads
        self.driver = None
        self.wait = None
        
        # Cria pasta para downloads se não existir
        if not os.path.exists(pasta_downloads):
            os.makedirs(pasta_downloads)
    
    def configurar_navegador(self):
        """Configura o navegador Chrome com as opções necessárias"""
        options = webdriver.ChromeOptions()
        
        # Configura pasta de download
        prefs = {
            "download.default_directory": os.path.abspath(self.pasta_downloads),
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True
        }
        options.add_experimental_option("prefs", prefs)
        
        # Mantém o navegador aberto para debug
        options.add_experimental_option("detach", True)
        
        self.driver = webdriver.Chrome(options=options)
        self.wait = WebDriverWait(self.driver, 10)
        
        print("✓ Navegador configurado com sucesso")
    
    def carregar_dados(self):
        """Carrega os dados do Excel"""
        try:
            df = pd.read_excel(self.caminho_excel)
            print(f"✓ Excel carregado: {len(df)} registros encontrados")
            
            # Adiciona colunas de controle se não existirem
            if 'Status' not in df.columns:
                df['Status'] = ''
            if 'Numero_Nota' not in df.columns:
                df['Numero_Nota'] = ''
            if 'Data_Emissao' not in df.columns:
                df['Data_Emissao'] = ''
            if 'Mensagem_Erro' not in df.columns:
                df['Mensagem_Erro'] = ''
            
            # Converte colunas para string para evitar erro de tipo
            df['Status'] = df['Status'].astype(str)
            df['Numero_Nota'] = df['Numero_Nota'].astype(str)
            df['Data_Emissao'] = df['Data_Emissao'].astype(str)
            df['Mensagem_Erro'] = df['Mensagem_Erro'].astype(str)
            
            return df
        except Exception as e:
            print(f"✗ Erro ao carregar Excel: {e}")
            return None
    
    def acessar_sistema(self):
        """Acessa o sistema da SEFIN (usuário já deve estar logado)"""
        url = "https://notafiscal.belem.pa.gov.br/notafiscal/paginas/notafiscal/emissaoNotaFiscalData.jsf?faces-redirect=true"
        
        print(f"  → Acessando: {url}")
        self.driver.get(url)
        
        print("  → Aguardando página carregar...")
        time.sleep(5)
        
        # Verifica se está na página correta
        print(f"  ℹ URL atual: {self.driver.current_url}")
        print(f"  ℹ Título da página: {self.driver.title}")
        
        # Salva screenshot para debug
        self.driver.save_screenshot("pagina_inicial.png")
        print(f"  ℹ Screenshot salvo: pagina_inicial.png")
        
        print("\n✓ Sistema acessado. Certifique-se de estar logado!")
        print("  ℹ Verifique se a página está correta antes de continuar")
        time.sleep(2)
    
    def preencher_cpf_e_pesquisar(self, cpf):
        """Preenche o CPF e clica em pesquisar"""
        try:
            # Remove formatação do CPF
            cpf_limpo = cpf.replace('.', '').replace('-', '').replace('/', '')
            print(f"  → Buscando campo CPF na seção Tomador...")
            
            # Aguarda a página carregar completamente
            time.sleep(2)
            
            # Lista todos os inputs da página para debug
            all_inputs = self.driver.find_elements(By.TAG_NAME, "input")
            print(f"  ℹ Total de campos input encontrados: {len(all_inputs)}")
            
            # Tenta encontrar o campo CPF de diferentes formas
            campo_cpf = None
            tentativas = [
                # Seletores corretos do sistema SEFIN Belém
                (By.ID, "formNotaFiscal:idCpfCnpjPessoa:idInputMaskCpfCnpj:inputText"),
                (By.NAME, "formNotaFiscal:idCpfCnpjPessoa:idInputMaskCpfCnpj:inputText"),
                (By.CSS_SELECTOR, "input.jarch-inputtext[data-p-label*='CPF']"),
                # Fallbacks
                (By.XPATH, "//input[contains(@id, 'idInputMaskCpfCnpj')]"),
                (By.XPATH, "//input[contains(@name, 'idInputMaskCpfCnpj')]"),
                (By.XPATH, "//label[contains(text(), 'CPF/CNPJ')]/following::input[1]"),
                (By.XPATH, "//input[@type='text' and contains(@id, 'CpfCnpj')]"),
            ]
            
            for i, (metodo, seletor) in enumerate(tentativas, 1):
                try:
                    campo_cpf = self.driver.find_element(metodo, seletor)
                    print(f"  ✓ Campo CPF encontrado (tentativa {i}): {seletor}")
                    break
                except:
                    continue
            
            if not campo_cpf:
                print(f"  ✗ Campo CPF não encontrado com nenhuma estratégia!")
                print(f"  ℹ URL atual: {self.driver.current_url}")
                
                # Lista os 5 primeiros inputs para debug
                print(f"  ℹ Primeiros inputs encontrados:")
                for i, inp in enumerate(all_inputs[:5], 1):
                    try:
                        print(f"    {i}. id='{inp.get_attribute('id')}' name='{inp.get_attribute('name')}' type='{inp.get_attribute('type')}'")
                    except:
                        pass
                
                self.driver.save_screenshot("erro_cpf.png")
                print(f"  ℹ Screenshot salvo: erro_cpf.png")
                return False
            
            # Preenche CPF
            campo_cpf.clear()
            time.sleep(0.5)
            campo_cpf.send_keys(cpf_limpo)
            print(f"  ✓ CPF {cpf} preenchido")
            time.sleep(1)
            
            # Clica em pesquisar
            print(f"  → Buscando botão Pesquisar...")
            btn_pesquisar = None
            tentativas_btn = [
                # Seletores corretos do sistema SEFIN Belém
                (By.CSS_SELECTOR, "a.btn.btn-success[onclick*='PrimeFaces']"),
                (By.XPATH, "//a[contains(@class, 'btn-success') and contains(@onclick, 'PrimeFaces')]"),
                (By.XPATH, "//a[contains(@id, 'idt') and contains(@class, 'btn-success')]"),
                # Fallbacks
                (By.XPATH, "//a[contains(@class, 'ui-commandlink') and contains(@class, 'btn-success')]"),
                (By.CSS_SELECTOR, "a.ui-commandlink.btn-success"),
                (By.XPATH, "//button[contains(., 'Pesquisar')]"),
                (By.XPATH, "//a[contains(@class, 'btn') and contains(., 'Pesquisar')]"),
            ]
            
            for i, (metodo, seletor) in enumerate(tentativas_btn, 1):
                try:
                    btn_pesquisar = self.driver.find_element(metodo, seletor)
                    print(f"  ✓ Botão Pesquisar encontrado (tentativa {i})")
                    break
                except:
                    continue
            
            if not btn_pesquisar:
                print(f"  ✗ Botão Pesquisar não encontrado!")
                
                # Lista botões para debug
                all_buttons = self.driver.find_elements(By.TAG_NAME, "button")
                print(f"  ℹ Total de botões encontrados: {len(all_buttons)}")
                print(f"  ℹ Primeiros botões:")
                for i, btn in enumerate(all_buttons[:5], 1):
                    try:
                        print(f"    {i}. Texto='{btn.text}' class='{btn.get_attribute('class')}'")
                    except:
                        pass
                
                self.driver.save_screenshot("erro_botao_pesquisar.png")
                print(f"  ℹ Screenshot salvo: erro_botao_pesquisar.png")
                return False
            
            btn_pesquisar.click()
            time.sleep(3)
            print(f"  ✓ Botão Pesquisar clicado com sucesso")
            return True
            
        except Exception as e:
            print(f"  ✗ Erro ao pesquisar CPF: {type(e).__name__}")
            print(f"  ℹ Detalhes: {str(e)[:200]}")
            self.driver.save_screenshot("erro_geral_cpf.png")
            print(f"  ℹ Screenshot salvo: erro_geral_cpf.png")
            return False
    
    def preencher_dados_cadastrais(self, dados):
        """Preenche os dados cadastrais do tomador"""
        try:
            print("  → Preenchendo dados cadastrais...")
            
            # Aguarda o formulário carregar após pesquisa
            time.sleep(2)
            
            # Nome/Razão Social
            if pd.notna(dados.get('Nome')):
                try:
                    campo_nome = self.driver.find_element(By.XPATH, "//input[contains(@id, 'nomeRazaoSocial') or contains(@name, 'nomeRazaoSocial')]")
                    campo_nome.clear()
                    campo_nome.send_keys(dados['Nome'])
                    print(f"    ✓ Nome preenchido")
                except:
                    print(f"    ⚠ Campo Nome não encontrado ou já preenchido")
            
            # Apelido (se existir)
            if pd.notna(dados.get('Apelido')):
                try:
                    campo_apelido = self.driver.find_element(By.XPATH, "//input[contains(@id, 'apelido') or contains(@name, 'apelido')]")
                    campo_apelido.clear()
                    campo_apelido.send_keys(dados['Apelido'])
                    print(f"    ✓ Apelido preenchido")
                except:
                    print(f"    ⚠ Campo Apelido não encontrado")
            
            # CEP
            if pd.notna(dados.get('CEP')):
                try:
                    campo_cep = self.driver.find_element(By.XPATH, "//input[contains(@id, 'cep') or contains(@name, 'cep')]")
                    campo_cep.clear()
                    cep_limpo = str(dados['CEP']).replace('-', '').replace('.', '')
                    campo_cep.send_keys(cep_limpo)
                    print(f"    ✓ CEP preenchido")
                    time.sleep(2)  # Aguarda carregar endereço
                except:
                    print(f"    ⚠ Campo CEP não encontrado")
            
            # Número
            if pd.notna(dados.get('Numero')):
                try:
                    campo_numero = self.driver.find_element(By.XPATH, "//input[contains(@id, 'numero') or contains(@name, 'numero')]")
                    campo_numero.clear()
                    campo_numero.send_keys(str(dados['Numero']))
                    print(f"    ✓ Número preenchido")
                except:
                    print(f"    ⚠ Campo Número não encontrado")
            
            # Complemento
            if pd.notna(dados.get('Complemento')) and str(dados['Complemento']).strip():
                try:
                    campo_complemento = self.driver.find_element(By.XPATH, "//input[contains(@id, 'complemento') or contains(@name, 'complemento')]")
                    campo_complemento.clear()
                    campo_complemento.send_keys(dados['Complemento'])
                    print(f"    ✓ Complemento preenchido")
                except:
                    print(f"    ⚠ Campo Complemento não encontrado")
            
            # Telefone
            if pd.notna(dados.get('Telefone')):
                try:
                    campo_telefone = self.driver.find_element(By.XPATH, "//input[contains(@id, 'telefone') or contains(@name, 'telefone')]")
                    campo_telefone.clear()
                    campo_telefone.send_keys(str(dados['Telefone']))
                    print(f"    ✓ Telefone preenchido")
                except:
                    print(f"    ⚠ Campo Telefone não encontrado")
            
            # Email
            if pd.notna(dados.get('Email')):
                try:
                    campo_email = self.driver.find_element(By.XPATH, "//input[contains(@id, 'email') or contains(@name, 'email')]")
                    campo_email.clear()
                    campo_email.send_keys(dados['Email'])
                    print(f"    ✓ Email preenchido")
                except:
                    print(f"    ⚠ Campo Email não encontrado")
            
            print("  ✓ Dados cadastrais processados")
            return True
            
        except Exception as e:
            print(f"  ✗ Erro ao preencher dados cadastrais: {type(e).__name__}")
            print(f"  ℹ Detalhes: {str(e)[:200]}")
            return False
    
    def selecionar_atividade(self):
        """Seleciona a atividade de condicionamento físico"""
        try:
            select_atividade = Select(self.driver.find_element(By.NAME, "formEmissaoNFe:atividadeEconomica"))
            select_atividade.select_by_value("931310000 - ATIVIDADES DE CONDICIONAMENTO FISICO")
            
            print("  ✓ Atividade econômica selecionada")
            return True
        except Exception as e:
            print(f"  ✗ Erro ao selecionar atividade: {e}")
            return False
    
    def adicionar_descricao_nota(self):
        """Adiciona a descrição favorita da nota fiscal"""
        try:
            # Clica no botão "Carregar descrição"
            btn_carregar = self.driver.find_element(By.XPATH, "//button[contains(text(), 'Carregar descrição')]")
            btn_carregar.click()
            
            time.sleep(1)
            
            # Seleciona a descrição "-" (primeira opção)
            checkbox_descricao = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//input[@type='checkbox' and contains(@id, 'checkboxDescricao')]"))
            )
            checkbox_descricao.click()
            
            # Clica em confirmar
            btn_confirmar = self.driver.find_element(By.XPATH, "//button[contains(@class, 'btn-success') and contains(text(), 'Confirmar')]")
            btn_confirmar.click()
            
            time.sleep(1)
            print("  ✓ Descrição da nota adicionada")
            return True
        except Exception as e:
            print(f"  ✗ Erro ao adicionar descrição: {e}")
            return False
    
    def preencher_valor(self, valor=110.00):
        """Preenche o valor dos serviços"""
        try:
            campo_valor = self.driver.find_element(By.NAME, "formEmissaoNFe:valorServicos")
            campo_valor.clear()
            campo_valor.send_keys(str(valor))
            
            # Pressiona Enter para calcular impostos
            campo_valor.send_keys(Keys.RETURN)
            
            time.sleep(2)
            print(f"  ✓ Valor R$ {valor:.2f} preenchido")
            return True
        except Exception as e:
            print(f"  ✗ Erro ao preencher valor: {e}")
            return False
    
    def emitir_nota(self):
        """Clica no botão para emitir a nota fiscal"""
        try:
            btn_emitir = self.driver.find_element(By.XPATH, "//button[contains(text(), 'Emitir Nota Fiscal')]")
            btn_emitir.click()
            
            time.sleep(3)
            
            # Tenta capturar o número da nota
            try:
                numero_nota = self.driver.find_element(By.XPATH, "//span[contains(@class, 'numero-nota')]").text
                print(f"  ✓ Nota fiscal emitida: {numero_nota}")
                return numero_nota
            except:
                print("  ✓ Nota fiscal emitida (número não capturado)")
                return "Emitida"
        except Exception as e:
            print(f"  ✗ Erro ao emitir nota: {e}")
            return None
    
    def fazer_download_pdf(self):
        """Faz o download do PDF da nota fiscal"""
        try:
            btn_download = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(@title, 'Download') or contains(text(), 'PDF')]"))
            )
            btn_download.click()
            
            time.sleep(3)
            print("  ✓ Download do PDF iniciado")
            return True
        except Exception as e:
            print(f"  ✗ Erro ao fazer download: {e}")
            return False
    
    def processar_nota(self, index, dados):
        """Processa uma única nota fiscal"""
        print(f"\n[{index + 1}] Processando nota para CPF: {dados['CPF']}")
        
        try:
            # 1. Preenche CPF e pesquisa
            if not self.preencher_cpf_e_pesquisar(dados['CPF']):
                return 'ERRO', '', 'Erro ao pesquisar CPF'
            
            # 2. Preenche dados cadastrais
            if not self.preencher_dados_cadastrais(dados):
                return 'ERRO', '', 'Erro ao preencher dados cadastrais'
            
            # 3. Seleciona atividade
            if not self.selecionar_atividade():
                return 'ERRO', '', 'Erro ao selecionar atividade'
            
            # 4. Adiciona descrição
            if not self.adicionar_descricao_nota():
                return 'ERRO', '', 'Erro ao adicionar descrição'
            
            # 5. Preenche valor
            valor = dados.get('Valor', 110.00)
            if not self.preencher_valor(valor):
                return 'ERRO', '', 'Erro ao preencher valor'
            
            # 6. Emite nota
            numero_nota = self.emitir_nota()
            if not numero_nota:
                return 'ERRO', '', 'Erro ao emitir nota'
            
            # 7. Faz download do PDF
            self.fazer_download_pdf()
            
            return 'EMITIDA', numero_nota, ''
            
        except Exception as e:
            return 'ERRO', '', str(e)
    
    def salvar_progresso(self, df):
        """Salva o progresso no Excel"""
        try:
            df.to_excel(self.caminho_excel, index=False)
            print("✓ Progresso salvo no Excel")
        except Exception as e:
            print(f"✗ Erro ao salvar progresso: {e}")
    
    def executar(self):
        """Executa o processo completo de automação"""
        print("=" * 60)
        print("SISTEMA DE AUTOMAÇÃO DE NOTAS FISCAIS - SEFIN BELÉM")
        print("=" * 60)
        
        # Carrega dados
        df = self.carregar_dados()
        if df is None:
            return
        
        # Configura navegador
        self.configurar_navegador()
        
        # Acessa sistema
        self.acessar_sistema()
        
        input("\n⚠ Certifique-se de estar LOGADO no sistema e pressione ENTER para continuar...")
        
        # DEBUG: Salva HTML da página
        print("\n[DEBUG] Salvando HTML da página para análise...")
        with open("pagina_debug.html", "w", encoding="utf-8") as f:
            f.write(self.driver.page_source)
        print("  ℹ HTML salvo em: pagina_debug.html")
        
        # Processa cada nota
        total = len(df)
        sucesso = 0
        erros = 0
        
        for index, row in df.iterrows():
            # Pula se já foi processada
            if row.get('Status') == 'EMITIDA':
                print(f"\n[{index + 1}/{total}] Nota já emitida para CPF {row['CPF']} - Pulando")
                sucesso += 1
                continue
            
            # Processa nota
            status, numero_nota, mensagem_erro = self.processar_nota(index, row)
            
            # Atualiza DataFrame
            df.at[index, 'Status'] = status
            df.at[index, 'Numero_Nota'] = numero_nota if numero_nota else ''
            df.at[index, 'Data_Emissao'] = datetime.now().strftime('%d/%m/%Y %H:%M:%S') if status == 'EMITIDA' else ''
            df.at[index, 'Mensagem_Erro'] = mensagem_erro if mensagem_erro else ''
            
            if status == 'EMITIDA':
                sucesso += 1
                print(f"  ✓ Sucesso! ({sucesso}/{total})")
            else:
                erros += 1
                print(f"  ✗ Erro! ({erros}/{total})")
            
            # Salva progresso a cada 5 notas
            if (index + 1) % 5 == 0:
                self.salvar_progresso(df)
            
            # Pausa entre requisições
            time.sleep(2)
        
        # Salva progresso final
        self.salvar_progresso(df)
        
        # Relatório final
        print("\n" + "=" * 60)
        print("RELATÓRIO FINAL")
        print("=" * 60)
        print(f"Total processado: {total}")
        print(f"✓ Sucesso: {sucesso}")
        print(f"✗ Erros: {erros}")
        print(f"Taxa de sucesso: {(sucesso/total)*100:.1f}%")
        print("=" * 60)
        
        input("\nPressione ENTER para fechar o navegador...")
        self.driver.quit()


# Função para criar planilha modelo
def criar_planilha_modelo(nome_arquivo="notas_fiscais_modelo.xlsx"):
    """Cria uma planilha Excel modelo com dados de exemplo"""
    
    dados_exemplo = {
        'CPF': [
            '020.892.442-67',
            '123.456.789-00',
            '987.654.321-00'
        ],
        'Nome': [
            'Adrielle da Silva Oliveira dos Santos',
            'João da Silva',
            'Maria Souza'
        ],
        'Apelido': [
            'ADRIELLE',
            'JOAO',
            'MARIA'
        ],
        'CEP': [
            '66113-000',
            '66010-000',
            '66020-000'
        ],
        'Complemento': [
            '',
            'Apto 101',
            'Casa'
        ],
        'Telefone': [
            '(91) 9982-28405',
            '(91) 98888-7777',
            '(91) 99999-8888'
        ],
        'Email': [
            'adriellesilvaoliveira.ao@gmail.com',
            'joao@email.com',
            'maria@email.com'
        ],
        'Valor': [
            110.00,
            110.00,
            110.00
        ]
    }
    
    df = pd.DataFrame(dados_exemplo)
    df.to_excel(nome_arquivo, index=False)
    print(f"✓ Planilha modelo criada: {nome_arquivo}")
    print(f"  Adicione mais linhas conforme necessário!")


# EXECUÇÃO PRINCIPAL
if __name__ == "__main__":
    import sys
    
    print("=" * 60)
    print("CONFIGURAÇÃO INICIAL")
    print("=" * 60)
    
    # Verifica se deve criar planilha modelo
    if len(sys.argv) > 1 and sys.argv[1] == "--criar-modelo":
        criar_planilha_modelo()
        sys.exit(0)
    
    # Solicita caminho do Excel
    caminho_excel = input("Digite o caminho do arquivo Excel (ou deixe em branco para usar 'notas_fiscais.xlsx'): ").strip()
    
    if not caminho_excel:
        caminho_excel = "notas_fiscais.xlsx"
    
    if not os.path.exists(caminho_excel):
        print(f"\n✗ Arquivo não encontrado: {caminho_excel}")
        print("\nDeseja criar uma planilha modelo? (s/n): ", end="")
        if input().lower() == 's':
            criar_planilha_modelo(caminho_excel)
            print("\n✓ Planilha criada! Preencha os dados e execute novamente.")
        sys.exit(1)
    
    # Executa automação
    automacao = AutomacaoNotaFiscal(caminho_excel)
    automacao.executar()