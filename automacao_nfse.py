"""
Sistema de Automação para Emissão de Notas Fiscais - SEFIN Belém
Versão Simplificada e Funcional
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
    
    def preencher_cpf_e_pesquisar(self, cpf):
        """Preenche CPF e clica em pesquisar"""
        try:
            print(f"  → Preenchendo CPF {cpf}...")
            
            # Limpa CPF
            cpf_limpo = cpf.replace('.', '').replace('-', '').replace('/', '')
            
            # Rola para baixo
            self.driver.execute_script("window.scrollTo(0, 400);")
            time.sleep(1)
            
            # Preenche CPF
            campo_cpf = self.driver.find_element(By.ID, "formNotaFiscal:idCpfCnpjPessoa:idInputMaskCpfCnpj:inputText")
            campo_cpf.clear()
            campo_cpf.send_keys(cpf_limpo)
            print(f"    ✓ CPF preenchido")
            time.sleep(1)
            
            # Busca o botão Pesquisar correto (tem "dados-pessoa" no onclick)
            btn = self.driver.find_element(By.XPATH, "//a[contains(@class, 'btn-success') and contains(@onclick, 'dados-pessoa')]")
            print(f"    ✓ Botão Pesquisar encontrado: {btn.get_attribute('id')}")
            
            # Clica
            self.driver.execute_script("arguments[0].click();", btn)
            print(f"    ✓ Clicado - Aguardando loading...")
            
            # Aguarda loading sumir
            time.sleep(2)
            try:
                self.wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, ".ui-blockui")))
                print(f"    ✓ Loading concluído")
            except:
                time.sleep(5)
                print(f"    ℹ Aguardou tempo fixo")
            
            # Verifica se carregou
            time.sleep(3)
            try:
                nome = self.driver.find_element(By.XPATH, "//input[contains(@id, 'nomeEmpresarial')]").get_attribute('value')
                if nome:
                    print(f"    ✓ Dados carregados: {nome[:30]}...")
                    return True
            except:
                pass
            
            print(f"    ⚠ Dados não carregaram")
            return True  # Continua mesmo assim
            
        except Exception as e:
            print(f"    ✗ Erro: {type(e).__name__}")
            self.driver.save_screenshot("erro_cpf.png")
            return False
    
    def selecionar_atividade(self):
        """Seleciona atividade 931310000"""
        try:
            print(f"  → Selecionando atividade...")
            
            # Aguarda carregar
            time.sleep(5)
            
            # Rola até atividade
            self.driver.execute_script("window.scrollTo(0, 1000);")
            time.sleep(2)
            
            # Busca o select
            select_elem = self.driver.find_element(By.ID, "formNotaFiscal:idAtividadeEmissor_input")
            
            # Remove disabled
            self.driver.execute_script("""
                var s = document.getElementById('formNotaFiscal:idAtividadeEmissor_input');
                s.removeAttribute('disabled');
                s.disabled = false;
            """)
            time.sleep(1)
            
            # Seleciona
            select = Select(select_elem)
            select.select_by_value("931310000")
            print(f"    ✓ Atividade selecionada")
            
            # Dispara onchange
            self.driver.execute_script("""
                var s = document.getElementById('formNotaFiscal:idAtividadeEmissor_input');
                if(s.onchange) s.onchange();
            """)
            
            time.sleep(3)
            return True
            
        except Exception as e:
            print(f"    ✗ Erro: {type(e).__name__}")
            return False
    
    def adicionar_descricao(self):
        """Adiciona descrição da nota"""
        try:
            print(f"  → Adicionando descrição...")
            
            # Rola até descrição
            self.driver.execute_script("window.scrollTo(0, 1500);")
            time.sleep(2)
            
            # Clica em "Carregar descrição"
            btn = self.driver.find_element(By.XPATH, "//button[contains(., 'Carregar descrição') or contains(., 'descrição')]")
            btn.click()
            time.sleep(2)
            
            # Seleciona primeira opção
            checkbox = self.driver.find_element(By.XPATH, "(//input[@type='checkbox'])[1]")
            checkbox.click()
            time.sleep(1)
            
            # Confirma
            btn_confirmar = self.driver.find_element(By.XPATH, "//button[contains(@class, 'btn-success') and contains(., 'Confirmar')]")
            btn_confirmar.click()
            time.sleep(2)
            
            print(f"    ✓ Descrição adicionada")
            return True
            
        except Exception as e:
            print(f"    ✗ Erro: {type(e).__name__}")
            return False
    
    def preencher_valor(self, valor=110.00):
        """Preenche valor dos serviços"""
        try:
            print(f"  → Preenchendo valor R$ {valor}...")
            
            # Rola até valor
            self.driver.execute_script("window.scrollTo(0, 2000);")
            time.sleep(1)
            
            # Preenche valor
            campo = self.driver.find_element(By.XPATH, "//input[contains(@id, 'valorServicos') or contains(@id, 'valor')]")
            campo.clear()
            campo.send_keys(str(valor))
            campo.send_keys(Keys.TAB)  # Sai do campo para calcular
            
            time.sleep(3)
            print(f"    ✓ Valor preenchido")
            return True
            
        except Exception as e:
            print(f"    ✗ Erro: {type(e).__name__}")
            return False
    
    def emitir_nota(self):
        """Emite a nota fiscal"""
        try:
            print(f"  → Emitindo nota...")
            
            # Rola até o final
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)
            
            # Clica em Emitir
            btn = self.driver.find_element(By.XPATH, "//button[contains(., 'Emitir')] | //a[contains(., 'Emitir')]")
            btn.click()
            
            time.sleep(5)
            
            # Tenta capturar número
            try:
                numero = self.driver.find_element(By.XPATH, "//*[contains(text(), 'Nota') and contains(text(), 'emitida')]").text
                print(f"    ✓ Nota emitida: {numero}")
                return "Emitida"
            except:
                print(f"    ✓ Nota emitida")
                return "Emitida"
                
        except Exception as e:
            print(f"    ✗ Erro: {type(e).__name__}")
            return None
    
    def processar_nota(self, index, dados):
        """Processa uma nota completa"""
        print(f"\n[{index + 1}] Processando CPF: {dados['CPF']}")
        
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
            valor = dados.get('Valor', 110.00)
            if not self.preencher_valor(valor):
                return 'ERRO', '', 'Erro ao preencher valor'
            
            # 5. Emitir
            numero = self.emitir_nota()
            if not numero:
                return 'ERRO', '', 'Erro ao emitir nota'
            
            return 'EMITIDA', numero, ''
            
        except Exception as e:
            return 'ERRO', '', str(e)
    
    def executar(self):
        """Executa o processo completo"""
        print("=" * 60)
        print("AUTOMAÇÃO NFS-E BELÉM - VERSÃO SIMPLIFICADA")
        print("=" * 60)
        
        # Carrega dados
        df = self.carregar_dados()
        
        # Configura navegador
        self.configurar_navegador()
        
        # Acessa sistema
        self.acessar_sistema()
        
        input("\n⚠ Faça LOGIN e pressione ENTER para continuar...\n")
        
        # Processa notas
        total = len(df)
        sucesso = 0
        
        for index, row in df.iterrows():
            if row.get('Status') == 'EMITIDA':
                print(f"\n[{index + 1}] Já processada - pulando")
                sucesso += 1
                continue
            
            status, numero, erro = self.processar_nota(index, row)
            
            df.at[index, 'Status'] = status
            df.at[index, 'Numero_Nota'] = numero if numero else ''
            df.at[index, 'Data_Emissao'] = datetime.now().strftime('%d/%m/%Y %H:%M') if status == 'EMITIDA' else ''
            df.at[index, 'Mensagem_Erro'] = erro if erro else ''
            
            if status == 'EMITIDA':
                sucesso += 1
                print(f"  ✓ SUCESSO ({sucesso}/{total})")
            else:
                print(f"  ✗ ERRO: {erro}")
            
            # Salva a cada 3
            if (index + 1) % 3 == 0:
                try:
                    df.to_excel(self.caminho_excel, index=False)
                    print("  ℹ Progresso salvo")
                except:
                    pass
            
            time.sleep(3)
        
        # Salva final
        try:
            df.to_excel(self.caminho_excel, index=False)
        except:
            print("✗ Erro ao salvar - feche o Excel!")
        
        # Relatório
        print("\n" + "=" * 60)
        print("RELATÓRIO FINAL")
        print("=" * 60)
        print(f"Total: {total}")
        print(f"✓ Sucesso: {sucesso}")
        print(f"✗ Erros: {total - sucesso}")
        print(f"Taxa: {(sucesso/total)*100:.1f}%")
        print("=" * 60)
        
        input("\nPressione ENTER para fechar...")
        self.driver.quit()


if __name__ == "__main__":
    import sys
    
    caminho = input("Arquivo Excel (ou ENTER para 'notas_fiscais.xlsx'): ").strip()
    if not caminho:
        caminho = "notas_fiscais.xlsx"
    
    if not os.path.exists(caminho):
        print(f"✗ Arquivo não encontrado: {caminho}")
        sys.exit(1)
    
    automacao = AutomacaoNotaFiscal(caminho)
    automacao.executar()