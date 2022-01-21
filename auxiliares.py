"""
Todas as funções de suporte
"""

import os
import sys
import time
import pypyodbc as pyodbc
import sensiveis as senha


def caminhoprojeto(subpasta=''):
    """
    :param subpasta: adiciona o caminho da subpasta dada como entrada (caso preenchido).
    :return: o caminho do projeto ao qual a rotina está inserida.
    """
    import errno

    try:
        if getattr(sys, 'frozen', False):
            caminho = os.path.dirname(sys.executable)
        else:
            caminho = os.path.dirname(os.path.abspath(__file__))

        if len(subpasta) > 0:
            if os.path.isdir(caminho + '\\' + subpasta):
                return caminho + '\\' + subpasta
            else:
                os.mkdir(caminho + '\\' + subpasta)

            if os.path.isdir(caminho + '\\' + subpasta):
                return caminho + '\\' + subpasta
            else:
                return ''
        else:
            if os.path.isdir(caminho):
                return caminho
            else:
                return ''

    except OSError as exc:
        if exc.errno != errno.EEXIST:
            raise
        pass


def ultimoarquivo(caminho, extensao):
    """
    :param caminho: diretório onde pesquisa será realizada.
    :param extensao: extensão do arquivo que está sendo buscado o último alterado.
    :return: retorna o caminho completo do último arquivo atualizado da extensão e caminho dado como entrada.
    """
    lista_arquivos = os.listdir(caminho)
    ultimadata = 0
    ultimoatualizado = ''
    for arquivo in lista_arquivos:
        if extensao.upper() in arquivo.upper():
            if os.path.getmtime(caminho + '/' + arquivo) > ultimadata:
                ultimoatualizado = caminho + '/' + arquivo

    return ultimoatualizado


def renomeararquivo(nomeantigo, novonome):
    """
    :param nomeantigo: nome do arquivo a ser renomeado (endereço completo).
    :param novonome: nome novo do arquivo (endereço completo).
    """
    if os.path.isfile(to_raw(novonome)):
        os.remove(to_raw(novonome))
    time.sleep(0.5)
    os.rename(to_raw(nomeantigo), to_raw(novonome))


def to_raw(string):
    """
    :param string: texto a ser tratado.
    :return: string com o prefixo f para criar strings literais formatadas e r usado para tornar a string numa
    string literal bruta e ignorar caracteres especiais 'dentro' dela como o '\', por exemplo.
    """
    return fr"{string}"


class Banco:
    """
    Criado para se conectar e realizar operações no banco de dados
    """

    def __init__(self, caminho):
        constr = 'Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=' + caminho + ';Pwd=' + senha.senhabanco
        conxn = pyodbc.connect(constr)
        self.cursor = conxn.cursor()

    def consultar(self, sql):
        """

        :param sql: código sql a ser executado (uma consulta SQL).
        :return: o resultado da consulta em uma lista.
        """
        self.cursor.execute(sql)
        resultado = self.cursor.fetchall()
        return resultado

    def fecharbanco(self):
        """
        Fecha a conexão com o banco de dados
        """
        self.cursor.close()


def quantidade_cores(caminho):
    """

    :param caminho: endereço da imagem salva.
    :return: a quantidade de cores presente na matriz de cores da imagem.
    """
    from PIL import Image
    # from matplotlib.pyplot import imshow

    if os.path.isfile(caminho):
        image = Image.open(caminho)
        # Corta a imagem para tirar as bordas inúteis
        image = image.crop((1, 0, 130, 24))
        # imshow(image)
        # Pega as cores e a quantidade de pixels (pelo que entendi) da mesma cor e coloca em uma lista
        # Calcula o número de pixels na tela (quant pixels eixo "x" * quant pixels eixo "y") para retornar
        # o máximo de cores possíveis (pior caso cada pixel de uma cor diferente), se tiver mais cores do que o definido
        # ao chamar o getcolors retorna None
        cores = image.convert('RGB').getcolors(image.size[0] * image.size[1])
        if cores is not None:
            # Retorna a quantidade de elementos da lista de cores
            return len(cores)
        else:
            return 0


def retornorvalorlista(lista, indice):
    """

    :param lista: lista a ser realizada a busca.
    :param indice: item a ser retornado.
    :return: item do índice dado como entrada.
    """
    import numpy as np

    lst = np.array(lista)
    return lst[indice]


def escreverlistaexcelog(caminho, lista):
    """

    :param caminho: caminho do arquivo a ser escrito.
    :param lista: lista a ser adicionada no arquivo do caminho dado.
    """
    import pandas as pd

    df = pd.DataFrame(lista)

    writer = pd.ExcelWriter(caminho, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Lista', index=False)
    writer.save()


def acertardataatual():
    """

    :return: retorna data e hora do sistema formatado.
    """
    from datetime import datetime

    textodata = datetime.now()
    return textodata.strftime('%Y_%m_%d_%H_%M_%S')


def caminhospadroes(caminho):
    """

    :param caminho: código dos caminhos padrões (dúvidas, confira lista abaixo).
    :return: retorna o caminho padrão selecionado (str)
    """
    import ctypes.wintypes
    # CSIDL	                        Decimal	Hex	    Shell	Description
    # CSIDL_ADMINTOOLS	            48	    0x30	5.0	    The file system directory that is used to store administrative tools for an individual user.
    # CSIDL_ALTSTARTUP	            29	    0x1D	 	    The file system directory that corresponds to the user's nonlocalized Startup program group.
    # CSIDL_APPDATA	                26	    0x1A	4.71	The file system directory that serves as a common repository for application-specific data.
    # CSIDL_BITBUCKET	            10	    0x0A	 	    The virtual folder containing the objects in the user's Recycle Bin.
    # CSIDL_CDBURN_AREA	            59	    0x3B	6.0	    The file system directory acting as a staging area for files waiting to be written to CD.
    # CSIDL_COMMON_ADMINTOOLS	    47	    0x2F	5.0	    The file system directory containing administrative tools for all users of the computer.
    # CSIDL_COMMON_ALTSTARTUP	    30	    0x1E	        NT-based only	The file system directory that corresponds to the nonlocalized Startup program group for all users.
    # CSIDL_COMMON_APPDATA	        35	    0x23	5.0	    The file system directory containing application data for all users.
    # CSIDL_COMMON_DESKTOPDIRECTORY	25	    0x19	        NT-based only	The file system directory that contains files and folders that appear on the desktop for all users.
    # CSIDL_COMMON_DOCUMENTS	    46	    0x2E	 	    The file system directory that contains documents that are common to all users.
    # CSIDL_COMMON_FAVORITES	    31	    0x1F	        NT-based only	The file system directory that serves as a common repository for favorite items common to all users.
    # CSIDL_COMMON_MUSIC	        53	    0x35	6.0	    The file system directory that serves as a repository for music files common to all users.
    # CSIDL_COMMON_PICTURES	        54	    0x36	6.0	    The file system directory that serves as a repository for image files common to all users.
    # CSIDL_COMMON_PROGRAMS	        23	    0x17	        NT-based only	The file system directory that contains the directories for the common program groups that appear on the Start menu for all users.
    # CSIDL_COMMON_STARTMENU	    22	    0x16	        NT-based only	The file system directory that contains the programs and folders that appear on the Start menu for all users.
    # CSIDL_COMMON_STARTUP	        24	    0x18	        NT-based only	The file system directory that contains the programs that appear in the Startup folder for all users.
    # CSIDL_COMMON_TEMPLATES	    45	    0x2D	        NT-based only	The file system directory that contains the templates that are available to all users.
    # CSIDL_COMMON_VIDEO	        55	    0x37	6.0	    The file system directory that serves as a repository for video files common to all users.
    # CSIDL_COMPUTERSNEARME	        61	    0x3D	6.0	    The folder representing other machines in your workgroup.
    # CSIDL_CONNECTIONS	            49	    0x31	6.0	    The virtual folder representing Network Connections, containing network and dial-up connections.
    # CSIDL_CONTROLS	            3	    0x03	 	    The virtual folder containing icons for the Control Panel applications.
    # CSIDL_COOKIES	                33	    0x21	 	    The file system directory that serves as a common repository for Internet cookies.
    # CSIDL_DESKTOP	                0	    0x00	 	    The virtual folder representing the Windows desktop, the root of the shell namespace.
    # CSIDL_DESKTOPDIRECTORY	    16	    0x10	 	    The file system directory used to physically store file objects on the desktop.
    # CSIDL_DRIVES	                17	    0x11	 	    The virtual folder representing My Computer, containing everything on the local computer: storage devices, printers, and Control Panel. The folder may also contain mapped network drives.
    # CSIDL_FAVORITES	            6	    0x06	 	    The file system directory that serves as a common repository for the user's favorite items.
    # CSIDL_FONTS	                20	    0x14	 	    A virtual folder containing fonts.
    # CSIDL_HISTORY	                34	    0x22	 	    The file system directory that serves as a common repository for Internet history items.
    # CSIDL_INTERNET	            1	    0x01	 	    A viritual folder for Internet Explorer.
    # CSIDL_INTERNET_CACHE	        32	    0x20	4.72	The file system directory that serves as a common repository for temporary Internet files.
    # CSIDL_LOCAL_APPDATA	        28	    0x1C	5.0	    The file system directory that serves as a data repository for local (nonroaming) applications.
    # CSIDL_MYDOCUMENTS	            5	    0x05	6.0	    The virtual folder representing the My Documents desktop item.
    # CSIDL_MYMUSIC	                13	    0x0D	6.0	    The file system directory that serves as a common repository for music files.
    # CSIDL_MYPICTURES	            39	    0x27	5.0	    The file system directory that serves as a common repository for image files.
    # CSIDL_MYVIDEO	                14	    0x0E	6.0	    The file system directory that serves as a common repository for video files.
    # CSIDL_NETHOOD	                19	    0x13	 	    A file system directory containing the link objects that may exist in the My Network Places virtual folder.
    # CSIDL_NETWORK	                18	    0x12	 	    A virtual folder representing Network Neighborhood, the root of the network namespace hierarchy.
    # CSIDL_PERSONAL	            5	    0x05	 	    The file system directory used to physically store a user's common repository of documents. (From shell version 6.0 onwards, CSIDL_PERSONAL is equivalent to CSIDL_MYDOCUMENTS, which is a virtual folder.)
    # CSIDL_PHOTOALBUMS	            69	    0x45	Vista	The virtual folder used to store photo albums.
    # CSIDL_PLAYLISTS	            63	    0x3F	Vista	The virtual folder used to store play albums.
    # CSIDL_PRINTERS	            4	    0x04	 	    The virtual folder containing installed printers.
    # CSIDL_PRINTHOOD	            27	    0x1B	 	    The file system directory that contains the link objects that can exist in the Printers virtual folder.
    # CSIDL_PROFILE	                40	    0x28	5.0	    The user's profile folder.
    # CSIDL_PROGRAM_FILES	        38	    0x26	5.0	    The Program Files folder.
    # CSIDL_PROGRAM_FILESX86	    42	    0x2A	5.0	    The Program Files folder for 32-bit programs on 64-bit systems.
    # CSIDL_PROGRAM_FILES_COMMON	43	    0x2B	5.0	    A folder for components that are shared across applications.
    # CSIDL_PROGRAM_FILES_COMMONX86	44	    0x2C	5.0	    A folder for 32-bit components that are shared across applications on 64-bit systems.
    # CSIDL_PROGRAMS	            2	    0x02	 	    The file system directory that contains the user's program groups (which are themselves file system directories).
    # CSIDL_RECENT	                8	    0x08	 	    The file system directory that contains shortcuts to the user's most recently used documents.
    # CSIDL_RESOURCES	            56	    0x38	6.0	    The file system directory that contains resource data.
    # CSIDL_RESOURCES_LOCALIZED	    57	    0x39	6.0	    The file system directory that contains localized resource data.
    # CSIDL_SAMPLE_MUSIC	        64	    0x40	Vista	The file system directory that contains sample music.
    # CSIDL_SAMPLE_PLAYLISTS	    65	    0x41	Vista	The file system directory that contains sample playlists.
    # CSIDL_SAMPLE_PICTURES	        66	    0x42	Vista	The file system directory that contains sample pictures.
    # CSIDL_SAMPLE_VIDEOS	        67	    0x43	Vista	The file system directory that contains sample videos.
    # CSIDL_SENDTO	                9	    0x09	 	    The file system directory that contains Send To menu items.
    # CSIDL_STARTMENU	            11	    0x0B	 	    The file system directory containing Start menu items.
    # CSIDL_STARTUP	                7	    0x07	 	    The file system directory that corresponds to the user's Startup program group.
    # CSIDL_SYSTEM	                37	    0x25	5.0	    The Windows System folder.
    # CSIDL_SYSTEMX86	            41	    0x29	5.0	    The Windows 32-bit System folder on 64-bit systems.
    # CSIDL_TEMPLATES	            21	    0x15	 	    The file system directory that serves as a common repository for document templates.
    # CSIDL_WINDOWS	                36	    0x24	5.0	    The Windows directory or SYSROOT.

    csidl_personal = caminho  # Caminho padrão
    shgfp_type_current = 0  # Para não pegar a pasta padrão e sim a definida como documentos

    buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
    ctypes.windll.Shell32.SHGetFolderPathW(None, csidl_personal, None, shgfp_type_current, buf)

    return buf.value


def caminhoselecionado(tipojanela=1, titulojanela='Selecione o caminho/arquivo:',
                       tipoarquivos=('Todos os Arquivos', '*.*'), caminhoini=caminhospadroes(5), arquivoinicial=''):
    """

    :param tipojanela: 1 — Seleciona Arquivo (Padrão); 2 — Seleciona caminho para salvar arquivo; 3 — Seleciona diretório.
    :param titulojanela: cabeçalho exibido na caixa de diálogo.
    :param tipoarquivos: extensão dos arquivos permitidos da seleção.
    :param caminhoini: caminho inicial.
    :param arquivoinicial: arquivo inicial.
    :return:
    """
    import tkinter as tk
    from tkinter import filedialog

    'Cria a janela raiz'
    root = tk.Tk()
    root.withdraw()

    if tipojanela == 1:
        retorno = filedialog.askopenfilename(title=titulojanela,
                                             initialdir=caminhoini, filetypes=tipoarquivos, initialfile=arquivoinicial)
        if retorno is None:  # asksaveasfile return `None` if dialog closed with "cancel".
            return None

    elif tipojanela == 2:
        name = filedialog.asksaveasfile(mode='w', defaultextension='.txt',
                                        filetypes=tipoarquivos,
                                        initialdir=caminhoini,
                                        title=titulojanela, initialfile=arquivoinicial)
        if name is None:  # asksaveasfile return `None` if dialog closed with "cancel".
            return None
        text2save = str(name.name)
        name.write('')
        retorno = text2save

    elif tipojanela == 3:
        name = filedialog.askdirectory(initialdir=caminhoini, title=titulojanela)
        if name is None:  # askdirectory return `None` if dialog closed with "cancel".
            return None
        text2save = name
        retorno = text2save

    else:
        return

    return retorno


def criarinputbox(titulo, mensagem, substituircaracter='', valorinicial=''):
    """
    :param valorinicial: valor pré-preenchido na caixa de texto.
    :param titulo: cabeçalho da caixa de recebimento de dados do usuário (inputbox).
    :param mensagem: mensagem (normalmente descritiva ao 'input') para orientar o usuário.
    :param substituircaracter: caso seja um campo de senha informar o parâmetro para que a digitação não fique visível.
    :return: janela com as opções escolhidas.
    """
    import tkinter as tk
    from tkinter import simpledialog

    root = tk.Tk()
    root.attributes("-topmost", True)
    root.withdraw()

    # the input dialog
    user_inp = simpledialog.askstring(title=titulo, prompt=mensagem, initialvalue=valorinicial, show=substituircaracter)
    if user_inp is None:
        user_inp = 0

    root.attributes("-topmost", True)

    return user_inp


def left(s, amount):
    """

    :param s: string.
    :param amount: quantidade de caracteres.
    :return: retorna à esquerda de uma string
    """
    return s[:amount]


def right(s, amount):
    """

    :param s: string.
    :param amount: quantidade de caracteres.
    :return: retorna à direita de uma string
    """
    return s[-amount:]


def mid(s, offset, amount):
    """

    :param s: string.
    :param offset: início do "corte" da string.
    :param amount: quantidade de caracteres.
    :return:
    """
    return s[(offset - 1):(offset - 1) + amount]
