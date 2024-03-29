from selenium import webdriver
import io
import sys
from PIL import Image
import os
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from anticaptchaofficial.imagecaptcha import *
from selenium.webdriver.support.ui import Select
import auxiliares as aux
from bs4 import BeautifulSoup
import messagebox
import sensiveis as senhas
from subprocess import CREATE_NO_WINDOW
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service


class TratarSite:
    """
    Classe que armazena todas as rotinas de execução de ações no controle remoto de site através do Python
    """

    def __init__(self, url, nomeperfil):
        self.url = url
        self.perfil = nomeperfil
        self.navegador = None
        self.options = None
        self.delay = 10
        self.caminho = ''

    def abrirnavegador(self):
        """
        :return: navegador configurado com o site desejado aberto
        """
        if self.navegador is not None:
            self.fecharsite()

        self.navegador = self.configuraprofilechrome()
        if self.navegador is not None:
            self.navegador.get(self.url)
            time.sleep(1)
            # Testa se a página carregou (ainda tem que fazer um teste e condição quando ele apresenta um texto de erro de carregamento)
            # ==========================================================================================================================
            resultadolimpo = ''
            corposite = BeautifulSoup(self.navegador.page_source, 'html.parser')
            for string in corposite.strings:
                resultadolimpo = resultadolimpo + ' ' + string

            if len(resultadolimpo) == 0:
                messagebox.msgbox(f'Site com problema de carregamento!', messagebox.MB_OK, 'Site fora do ar')
                self.navegador = -1
            # ==========================================================================================================================
            time.sleep(1)
            return self.navegador

    def configuraprofilechrome(self):
        """
        Configura usuário e opções no navegador aberto para execução
        return: o navegador configurado para iniciar a execução das rotinas
        """
        self.options = webdriver.ChromeOptions()
        if aux.caminhoprojeto('Profile') != '':
            self.options.add_argument("user-data-dir=" + aux.caminhoprojeto('Profile'))
            self.options.add_argument("--start-maximized")
            self.options.add_argument("---printing")
            self.options.add_argument("--disable-print-preview")
            self.options.add_argument("--silent")
            # Forma invisível
            # self.options.add_argument("--headless")

            if aux.caminhoprojeto('Downloads') != '':
                self.options.add_experimental_option('prefs', {
                    "profile.name": self.perfil,
                    "download.default_directory": aux.caminhoprojeto('Downloads'),  # Change default directory for downloads
                    "download.prompt_for_download": False,  # To auto download the file
                    "download.directory_upgrade": True,
                    "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
                })

        servico = Service(ChromeDriverManager().install())
        servico.creationflags = CREATE_NO_WINDOW

        return webdriver.Chrome(service=servico, options=self.options)

    def verificarobjetoexiste(self, identificador, endereco, valorselecao='', itemunico=True, iraoobjeto=False):
        """

        :param iraoobjeto: se simula o mouse em cima do objeto ou não.
        :param identificador: como será identificado, por nome, por nome de classe, etc.
        :param endereco: nome do objeto no site (lembrando que o nome é segundo o parâmetro anterior, se for definido ID no parâmetro anterior,
                         nesse tem que vir o ID do objeto do site, por exemplo.
        :param valorselecao: caso se um combobox passar o valor de seleção desejado que ele mesmo seleciona o dado nesse parâmetro.
        :param itemunico: caso seja uma coleção de objetos que queira selecionar (todos o do nome de classe x), colocar False, padrão será item único.
        :return: vai retornar o objeto do site para ser trabalhado já verificando se o mesmo existe, caso não encontre retorna None
        """
        if self.navegador is not None:
            try:
                if itemunico:
                    if len(valorselecao) == 0:
                        elemento = WebDriverWait(self.navegador, self.delay).until(EC.visibility_of_element_located((getattr(By, identificador), endereco)))
                        # elemento = self.navegador.find_element(getattr(By, identificador), endereco)

                    else:
                        # elemento = WebDriverWait(self.navegador, self.delay).until(EC.visibility_of_element_located((getattr(By, identificador), endereco)))
                        elemento = Select(self.navegador.find_element(getattr(By, identificador), endereco))
                        elemento.select_by_visible_text(valorselecao)
                else:
                    if len(valorselecao) == 0:
                        elemento = self.navegador.find_elements(getattr(By, identificador), endereco)
                    else:
                        elemento = Select(self.navegador.find_elements(getattr(By, identificador), endereco))
                        elemento.select_by_value(valorselecao)

                if iraoobjeto and elemento is not None:
                    action = ActionChains(self.navegador)
                    # if getattr(sys, 'frozen', False):
                    action.move_to_element(elemento).click().perform()
                    # else:
                    #    self.navegador.execute_script("arguments[0].click()", elemento)

                return elemento

            except NoSuchElementException:
                return None

            except TimeoutException:
                # messagebox.msgbox('Erro de carregamento objeto!', messagebox.MB_OK, 'Erro Carregamento')
                return None

    def descerrolagem(self):
        """
        Desce a barra de rolagem do navegador
        """
        self.navegador.execute_script("window.scrollTo(0,document.body.scrollHeight)")
        time.sleep(1)

    def mexerzoom(self, valor):
        """
        Desce a barra de rolagem do navegador
        """
        self.navegador.execute_script("document.body.style.transform='scale(" + str(valor) + ")';")
        time.sleep(1)

    def baixarimagem(self, identificador, endereco, caminho):
        """

        :param identificador: como será identificado, por nome, por nome de classe, etc.
        :param endereco: nome do objeto no site (lembrando que o nome é segundo o parâmetro anterior, se for definido ID no parâmetro anterior,
                         nesse tem que vir o ID do objeto do site, por exemplo.
        :param caminho: caminho onde a imagem será salva.
        :return: informa se achou a imagem no site e se conseguiu salvar a mesma.
        """
        achouimagem = False
        salvouimagem = False
        if self.navegador is not None:
            if os.path.isfile(caminho):
                os.remove(caminho)
            elemento = self.verificarobjetoexiste(identificador, endereco)
            if elemento is not None:
                achouimagem = True
                image = elemento.screenshot_as_png
                imagestream = io.BytesIO(image)
                im = Image.open(imagestream)
                im.save(caminho)
                if os.path.isfile(caminho):
                    # Verifica se a imagem veio toda "preta" (quando o site não carrega o CAPTCHA)
                    if aux.quantidade_cores(caminho) > 1:
                        salvouimagem = True
                        self.caminho = caminho
                    else:
                        os.remove(caminho)
                        salvouimagem = False
                else:
                    salvouimagem = False

        return achouimagem, salvouimagem

    @staticmethod
    def esperadownloads(caminho, timeout, nfiles=None):
        """
        Wait for downloads to finish with a specified timeout.
        Args
        ----
        caminho : str
            The path to the folder where the files will be downloaded.
        timeout : int
            How many seconds to wait until timing out.
        nfiles : int, defaults to None
            If provided, also wait for the expected number of files.
        """
        time.sleep(1)
        seconds = 0
        dl_wait = True
        while dl_wait and seconds < timeout:
            time.sleep(1)
            dl_wait = False
            files = os.listdir(caminho)
            if nfiles and len(files) != nfiles:
                dl_wait = True

            for fname in files:
                if fname.endswith('.crdownload'):
                    dl_wait = True

            seconds += 1

        time.sleep(1)
        return seconds

    def num_abas(self):
        """
        :return: a quantidade de abas abertas no navegador
        """
        if self.navegador is not None:
            return len(self.navegador.window_handles)

    def fecharsite(self):
        """
        fecha o browser carregado no objeto
        """
        print(hasattr(self.navegador, 'quit'))
        if self.navegador is not None and hasattr(self.navegador, 'quit'):
            self.navegador.quit()

    def resolvercaptcha(self, identificacaocaixa, caixacaptcha, identicacaobotao, botao):
        """
        :param identificacaocaixa: opção de como a caixa de texto do captcha será identificada (ID, NAME, CLASS, ETC.)
        :param caixacaptcha: idenficação da caixa do captcha segundo a variável anterior (se for ID, colocar o nome ID, por exemplo)
        :param identicacaobotao: opção de como o botão de ação do captcha será identificado (ID, NAME, CLASS, ETC.)
        :param botao: idenficação do botão de ação do captcha segundo a variável anterior (se for ID, colocar o nome ID, por exemplo)
        :return: booleana dizendo se conseguiu ou não resolver o captcha
        """
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as ec
        from selenium.common.exceptions import TimeoutException

        solver = imagecaptcha()
        solver.set_verbose(1)
        solver.set_key(senhas.chaveanticaptcha)

        captcha_text = solver.solve_and_return_solution(self.caminho)

        if len(str(captcha_text)) != 0:
            captcha = self.verificarobjetoexiste(identificacaocaixa, caixacaptcha)
            captcha.click()
            captcha.send_keys(captcha_text)
            time.sleep(2)
            confirmar = self.verificarobjetoexiste(identicacaobotao, botao, iraoobjeto=True)
            time.sleep(2)

            # time.sleep(2)
            try:
                alert = WebDriverWait(self.navegador, 2).until(ec.alert_is_present())
                texto = alert.text
                texto = texto.strip()
                alert.accept()
                if texto == 'CÓDIGO digitado Inválido!':
                    solver.report_incorrect_image_captcha()
                    resposta = False
                else:
                    resposta = True
                    if os.path.isfile(self.caminho):
                        os.remove(self.caminho)
                        self.caminho = ''

            except TimeoutException:
                resposta = True
                if os.path.isfile(self.caminho):
                    os.remove(self.caminho)
                    self.caminho = ''

        else:
            print("Erro solução captcha:" + solver.error_code)
            resposta = False

        return resposta
