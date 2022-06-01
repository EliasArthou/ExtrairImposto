"""
Extração de IPTUs
"""


def extrairboletos(visual):
    """

    : param caminhobanco: caminho do banco para realizar a pesquisa.
    : param resposta: opção selecionada de extração.
    : param visual: janela a ser manipulada.
    """

    import os
    import web
    import auxiliares as aux
    import sensiveis as senha
    import time
    import messagebox
    import sys

    codigocliente = ''
    site = None
    listaexcel = []

    # try:
    gerarboleto = not visual.var1.get()
    salvardadospdf = visual.var2.get()

    visual.acertaconfjanela(False)

    if os.path.isfile(aux.caminhoprojeto()+'/'+'Scai.WMB'):
        caminhobanco = aux.caminhoselecionado(titulojanela='Selecione o arquivo de banco de dados:',
                                              tipoarquivos=[('Banco ' + senha.empresa, '*.WMB'), ('Todos os Arquivos:', '*.*')],
                                              caminhoini=aux.caminhoprojeto(), arquivoinicial='Scai.WMB')
    else:
        if os.path.isdir(aux.caminhoprojeto()):
            caminhobanco = aux.caminhoselecionado(titulojanela='Selecione o arquivo de banco de dados:',
                                                  tipoarquivos=[('Banco ' + senha.empresa, '*.WMB'), ('Todos os Arquivos:', '*.*')],
                                                  caminhoini=aux.caminhoprojeto())
        else:
            caminhobanco = aux.caminhoselecionado(titulojanela='Selecione o arquivo de banco de dados:',
                                                  tipoarquivos=[('Banco ' + senha.empresa, '*.WMB'), ('Todos os Arquivos:', '*.*')])

    if len(caminhobanco) == 0:
        messagebox.msgbox('Selecione o caminho do Banco de Dados!', messagebox.MB_OK, 'Erro Banco')
        visual.manipularradio(True)
        sys.exit()

    resposta = str(visual.radio_valor.get())
    indicecliente = aux.criarinputbox('Cliente de Corte', 'Iniciar a partir de um cliente? (0 fará de todos da lista)', valorinicial='0')

    if indicecliente is not None:
        if not indicecliente.isdigit():
            messagebox.msgbox('Digite um valor válido (precisa ser numérico)!', messagebox.MB_OK, 'Opção Inválida')
            visual.manipularradio(True)
            sys.exit()
    else:
        messagebox.msgbox('Digite o inicío desejado ou deixe 0 (Zero)!', messagebox.MB_OK, 'Opção Inválida!')
        visual.manipularradio(True)
        sys.exit()

    tempoinicio = time.time()

    visual.acertaconfjanela(True)

    visual.mudartexto('labelstatus', 'Executando pesquisa no banco...')

    bd = aux.Banco(caminhobanco)
    indicecliente = str(indicecliente).zfill(4)
    if indicecliente == '0000':
        resultado = bd.consultar("SELECT * FROM [Lista Codigos IPTUs Completa]")
    else:
        resultado = bd.consultar("SELECT * FROM [Lista Codigos IPTUs Completa] WHERE Codigo >= '{codigo}'".format(codigo=indicecliente))

    bd.fecharbanco()

    pastadownload = aux.caminhoprojeto() + '\\' + 'Downloads'
    listachaves = ['Código Cliente', 'Inscrição', 'Guia do Exercício', 'Nr Guia', 'Valor', 'Contribuinte', 'Endereço', 'Status']
    listaexcel = []
    site = web.TratarSite(senha.site, 'ExtrairBoletoIPTU')

    for indice, linha in enumerate(resultado):
        codigocliente = linha['codigo']
        # ===================================== Parte Gráfica =======================================================
        visual.mudartexto('labelcodigocliente', 'Código Cliente: ' + codigocliente)
        visual.mudartexto('labelinscricao', str(linha['iptu']))
        visual.mudartexto('labelquantidade', 'Item ' + str(indice + 1) + ' de ' + str(len(resultado)) + '...')
        visual.mudartexto('labelstatus', 'Extraindo boleto...')
        # Atualiza a barra de progresso das transações (Views)
        visual.configurarbarra('barraextracao', len(resultado), indice + 1)
        time.sleep(0.1)
        # ===================================== Parte Gráfica =======================================================
        caminhodestino = pastadownload + '/' + codigocliente + '_' + linha['iptu'] + '.pdf'
        if not os.path.isfile(caminhodestino) or not gerarboleto:
            if site is not None:
                site.fecharsite()
            site = web.TratarSite(senha.site, 'ExtrairBoletoIPTU')
            site.abrirnavegador()
            if site.url != senha.site or site is None:
                if site is not None:
                    site.fecharsite()
                site = web.TratarSite(senha.site, senha.nomeprofile)
                site.abrirnavegador()

            if site is not None and site.navegador != -1:
                # Campo de Inscrição da tela Inicial
                inscricao = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_inscricao_input')
                if inscricao is not None:
                    inscricao.clear()
                    inscricao.send_keys(linha['iptu'])
                    if site.url == senha.site:
                        botaogerar = site.verificarobjetoexiste('NAME', 'ctl00$ePortalContent$DefiniGuia')
                        if botaogerar is not None:
                            if getattr(sys, 'frozen', False):
                                botaogerar.click()
                            else:
                                site.navegador.execute_script("arguments[0].click()", botaogerar)

                            mensagemerro = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_MSG')
                            if mensagemerro is None:
                                site.delay = 2
                                mensagemguia = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_TELA_M1')
                                site.delay = 10
                                if mensagemguia is None:
                                    guia = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_TELA_Guia1')
                                    if guia is not None:
                                        if getattr(sys, 'frozen', False):
                                            guia.click()
                                        else:
                                            site.navegador.execute_script("arguments[0].click()", guia)
                                        guiaexercicio = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_GuiaExercicio')
                                        if guiaexercicio is None:
                                            guiaexercicio = ''
                                        else:
                                            guiaexercicio = guiaexercicio.text

                                        contribuinte = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_TELA_CONTRIBUINTE')
                                        if contribuinte is None:
                                            contribuinte = ''
                                        else:
                                            contribuinte = contribuinte.text

                                        endereco = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_TELA_ENDERECO')
                                        if endereco is None:
                                            endereco = ''
                                        else:
                                            endereco = endereco.text

                                        botaogerarid = 'ctl00_ePortalContent_btnDarmIndiv'

                                        match resposta:
                                            case '1':
                                                idselecionado = 'ctl00_ePortalContent_cbCotaUnica'
                                                namevalor = 'ctl00$ePortalContent$cbCotaUnica'
                                                valores = site.verificarobjetoexiste('NAME', namevalor)
                                                dadosiptu = [codigocliente, linha['iptu'], guiaexercicio, 1, valores.text, contribuinte, endereco]
                                                listaexcel.append(dict(zip(listachaves, dadosiptu)))

                                            case '2' | '3' | '4':
                                                idselecionado = 'ctl00_ePortalContent_Chk_00' + str(int(resposta) - 1)
                                                if resposta == '2':
                                                    namevalor = 'Valor_Prim'
                                                elif resposta == '3':
                                                    namevalor = 'Valor_Segu'
                                                else:
                                                    namevalor = 'Valor_Terc'

                                                valores = site.verificarobjetoexiste('NAME', namevalor, itemunico=False)
                                                for index, valor in enumerate(valores):
                                                    dadosiptu = [codigocliente, linha['iptu'], guiaexercicio, str(index + 1), valor.text,
                                                                 contribuinte, endereco]
                                                    listaexcel.append(dict(zip(listachaves, dadosiptu)))

                                            case _:
                                                idselecionado = ''

                                        if gerarboleto:
                                            # site.mexerzoom(0.5)
                                            cota = site.verificarobjetoexiste('ID', idselecionado, iraoobjeto=True)
                                            if cota is not None:
                                                # site.descerrolagem()
                                                botaogerar = site.verificarobjetoexiste('ID', botaogerarid, iraoobjeto=True)
                                                if botaogerar is not None:
                                                    confirmar = site.verificarobjetoexiste('ID', 'popup_ok')
                                                    if confirmar is not None:
                                                        if getattr(sys, 'frozen', False):
                                                            confirmar.click()
                                                        else:
                                                            site.navegador.execute_script("arguments[0].click()", confirmar)

                                                        if site.navegador.current_url == senha.telaboleto:
                                                            imprimir = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_ImprimirDARM',
                                                                                                  iraoobjeto=True)
                                                            if imprimir is not None:
                                                                visual.mudartexto('labelstatus', 'Salvando Boleto...')
                                                                site.esperadownloads(pastadownload, 10)
                                                                baixado = aux.ultimoarquivo(pastadownload, '.pdf')
                                                                if 'DARM_' not in baixado:
                                                                    baixado = ''
                                                                if len(baixado) > 0:
                                                                    caminhodestino = aux.to_raw(caminhodestino)
                                                                    aux.adicionarcabecalhopdf(baixado, caminhodestino, codigocliente)
                                                                    if salvardadospdf:
                                                                        listacodigo = []
                                                                        listatipopag = []
                                                                        listaarquivo = []
                                                                        df = aux.extrairtextopdf(caminhodestino)
                                                                        total_rows = df[df.columns[0]].count()
                                                                        for linha in range(total_rows):
                                                                            listacodigo.append(codigocliente)
                                                                            listatipopag.append('PARCELADO')
                                                                            listaarquivo.append(caminhodestino)

                                                                        df.insert(loc=0, column='Codigo', value=listacodigo)
                                                                        df.insert(loc=4, column='TpoPagto', value=listatipopag)
                                                                        df.insert(loc=5, column='Arquivo', value=listaarquivo)

                                                                        bd.adicionardf('Codigos IPTUs', df, 8)
                                                                    # aux.renomeararquivo(baixado, pastadownload + '/' + codigocliente + '_' + linha['iptu'] + '.pdf')

                                        if os.path.isfile(pastadownload + '/' + codigocliente + '_' + linha['iptu'] + '.pdf'):
                                            for dicionario in listaexcel:

                                                if dicionario['Código Cliente'] == codigocliente and dicionario['Inscrição'] == linha['iptu']:
                                                    dicionario.update({'Status': 'Ok'})
                                        else:
                                            for dicionario in listaexcel:
                                                if dicionario['Código Cliente'] == codigocliente and dicionario['Inscrição'] == linha['iptu']:
                                                    dicionario.update({'Status': 'Verificar'})

                                        if site is not None:
                                            site.fecharsite()
                                else:
                                    valorimpostotela = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_TELA_Valor1')
                                    if valorimpostotela is not None:
                                        valor = valorimpostotela.text
                                        valor = valor.replace('.', '')
                                        valor = valor.replace(',', '.')
                                    else:
                                        valor = 0

                                    guiaexercicio = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_TELA_Guia1')
                                    if guiaexercicio is None:
                                        guiaexercicio = ''
                                    else:
                                        guiaexercicio = guiaexercicio.text

                                    contribuinte = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_TELA_CONTRIBUINTE')
                                    if contribuinte is None:
                                        contribuinte = ''
                                    else:
                                        contribuinte = contribuinte.text

                                    endereco = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_TELA_ENDERECO')
                                    if endereco is None:
                                        endereco = ''
                                    else:
                                        endereco = endereco.text

                                    if valorimpostotela is None or float(valor) == 0:
                                        dadosiptu = [codigocliente, linha['iptu'], guiaexercicio, '0', '0,00', contribuinte, endereco, 'Sem Guia (Provável Isento)']
                                    else:
                                        dadosiptu = [codigocliente, linha['iptu'], guiaexercicio, '0', valorimpostotela.text, contribuinte, endereco,
                                                     'Verificar (Extrair Manualmente)']

                                    listaexcel.append(dict(zip(listachaves, dadosiptu)))
                                    if site is not None:
                                        site.fecharsite()

                            else:
                                dadosiptu = [codigocliente, linha['iptu'], '', '', '', '', '', mensagemerro.text]
                                listaexcel.append(dict(zip(listachaves, dadosiptu)))
                                site.fecharsite()
        else:
            if os.path.isfile(caminhodestino) and salvardadospdf:
                listacodigo = []
                listatipopag = []
                listaarquivo = []
                df = aux.extrairtextopdf(caminhodestino)
                total_rows = df[df.columns[0]].count()
                for linha in range(total_rows):
                    listacodigo.append(codigocliente)
                    listatipopag.append('PARCELADO')
                    listaarquivo.append(caminhodestino)

                df.insert(loc=0, column='Codigo', value=listacodigo)
                df.insert(loc=4, column='TpoPagto', value=listatipopag)
                df.insert(loc=5, column='Arquivo', value=listaarquivo)

                bd.adicionardf('Codigos IPTUs', df, 7)
    # finally:
    print(codigocliente)
    if site is not None:
        site.fecharsite()

    if len(listaexcel) > 0:
        visual.mudartexto('labelstatus', 'Salvando lista...')
        aux.escreverlistaexcelog('Log_' + aux.acertardataatual() + '.xlsx', listaexcel)

    tempofim = time.time()

    hours, rem = divmod(tempofim-tempoinicio, 3600)
    minutes, seconds = divmod(rem, 60)

    visual.manipularradio(True)
    visual.acertaconfjanela(False)

    messagebox.msgbox(
        f'O tempo decorrido foi de: {"{:0>2}:{:0>2}:{:05.2f}".format(int(hours), int(minutes), int(seconds))}',
        messagebox.MB_OK, 'Tempo Decorrido')


def extrairnadaconsta(visual):
    """

    : param caminhobanco: caminho do banco para realizar a pesquisa.
    : param resposta: opção selecionada de extração.
    : param visual: janela a ser manipulada.
    """

    import os
    import web
    import auxiliares as aux
    import sensiveis as senha
    import time
    import messagebox
    import sys

    codigocliente = ''
    site = None
    listaexcel = []

    # try:
    gerarboleto = not visual.var1.get()
    salvardadospdf = visual.var2.get()

    visual.acertaconfjanela(False)

    if os.path.isfile(aux.caminhoprojeto()+'/'+'Scai.WMB'):
        caminhobanco = aux.caminhoselecionado(titulojanela='Selecione o arquivo de banco de dados:',
                                              tipoarquivos=[('Banco ' + senha.empresa, '*.WMB'), ('Todos os Arquivos:', '*.*')],
                                              caminhoini=aux.caminhoprojeto(), arquivoinicial='Scai.WMB')
    else:
        if os.path.isdir(aux.caminhoprojeto()):
            caminhobanco = aux.caminhoselecionado(titulojanela='Selecione o arquivo de banco de dados:',
                                                  tipoarquivos=[('Banco ' + senha.empresa, '*.WMB'), ('Todos os Arquivos:', '*.*')],
                                                  caminhoini=aux.caminhoprojeto())
        else:
            caminhobanco = aux.caminhoselecionado(titulojanela='Selecione o arquivo de banco de dados:',
                                                  tipoarquivos=[('Banco ' + senha.empresa, '*.WMB'), ('Todos os Arquivos:', '*.*')])

    if len(caminhobanco) == 0:
        messagebox.msgbox('Selecione o caminho do Banco de Dados!', messagebox.MB_OK, 'Erro Banco')
        visual.manipularradio(True)
        sys.exit()

    resposta = str(visual.radio_valor.get())
    indicecliente = aux.criarinputbox('Cliente de Corte', 'Iniciar a partir de um cliente? (0 fará de todos da lista)', valorinicial='0')

    if indicecliente is not None:
        if not indicecliente.isdigit():
            messagebox.msgbox('Digite um valor válido (precisa ser numérico)!', messagebox.MB_OK, 'Opção Inválida')
            visual.manipularradio(True)
            sys.exit()
    else:
        messagebox.msgbox('Digite o inicío desejado ou deixe 0 (Zero)!', messagebox.MB_OK, 'Opção Inválida!')
        visual.manipularradio(True)
        sys.exit()

    tempoinicio = time.time()

    visual.acertaconfjanela(True)

    visual.mudartexto('labelstatus', 'Executando pesquisa no banco...')

    bd = aux.Banco(caminhobanco)
    indicecliente = str(indicecliente).zfill(4)
    if indicecliente == '0000':
        resultado = bd.consultar("SELECT * FROM [Lista Codigos IPTUs Completa]")
    else:
        resultado = bd.consultar("SELECT * FROM [Lista Codigos IPTUs Completa] WHERE Codigo >= '{codigo}'".format(codigo=indicecliente))

    bd.fecharbanco()

    pastadownload = aux.caminhoprojeto() + '\\' + 'Downloads'
    listachaves = ['Código Cliente', 'Inscrição', 'Guia do Exercício', 'Status']
    listaexcel = []
    site = web.TratarSite(senha.site, 'ExtrairBoletoIPTU')

    for indice, linha in enumerate(resultado):
        codigocliente = linha['codigo']
        # ===================================== Parte Gráfica =======================================================
        visual.mudartexto('labelcodigocliente', 'Código Cliente: ' + codigocliente)
        visual.mudartexto('labelinscricao', str(linha['iptu']))
        visual.mudartexto('labelquantidade', 'Item ' + str(indice + 1) + ' de ' + str(len(resultado)) + '...')
        visual.mudartexto('labelstatus', 'Extraindo boleto...')
        # Atualiza a barra de progresso das transações (Views)
        visual.configurarbarra('barraextracao', len(resultado), indice + 1)
        time.sleep(0.1)
        # ===================================== Parte Gráfica =======================================================
        caminhodestino = pastadownload + '/' + codigocliente + '_' + linha['iptu'] + '.pdf'
        if not os.path.isfile(caminhodestino) or not gerarboleto:
            if site is not None:
                site.fecharsite()
            site = web.TratarSite(senha.sitenadaconsta, 'ExtrairBoletoIPTU')
            site.abrirnavegador()
            if site.url != senha.sitenadaconsta or site is None:
                if site is not None:
                    site.fecharsite()
                site = web.TratarSite(senha.site, senha.nomeprofile)
                site.abrirnavegador()

            if site is not None and site.navegador != -1:
                # Campo de Inscrição da tela Inicial
                inscricao = site.verificarobjetoexiste('ID', 'Inscricao')
                if inscricao is not None:
                    inscricao.clear()
                    inscricao.send_keys(linha['iptu'])
                    if site.url == senha.sitenadaconsta:
                        botaogerar = site.verificarobjetoexiste('ID', 'Avancar')
                        if botaogerar is not None:
                            if getattr(sys, 'frozen', False):
                                botaogerar.click()
                            else:
                                site.navegador.execute_script("arguments[0].click()", botaogerar)

                            site.delay = 2
                            mensagemerro = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_MSG')
                            site.delay = 10
                            if mensagemerro is None:
                                linkdownload = site.verificarobjetoexiste('LINK_TEXT', 'link')
                                if linkdownload is not None:
                                    if botaogerar is not None:
                                        if getattr(sys, 'frozen', False):
                                            linkdownload.click()
                                        else:
                                            site.navegador.execute_script("arguments[0].click()", linkdownload)

                                    dadosiptu = [codigocliente, linha['iptu'], '2022']
                                    listaexcel.append(dict(zip(listachaves, dadosiptu)))
                                    site.esperadownloads(pastadownload, 10)
                                    baixado = aux.ultimoarquivo(pastadownload, '.pdf')
                                    if 'DARM_' not in baixado:
                                        baixado = ''
                                    if len(baixado) > 0:
                                        caminhodestino = aux.to_raw(caminhodestino)
                                        aux.adicionarcabecalhopdf(baixado, caminhodestino, codigocliente)

                                    if os.path.isfile(pastadownload + '/' + codigocliente + '_' + linha['iptu'] + '.pdf'):
                                        for dicionario in listaexcel:
                                            if dicionario['Código Cliente'] == codigocliente and dicionario['Inscrição'] == linha['iptu']:
                                                dicionario.update({'Status': 'Ok'})
                                    else:
                                        for dicionario in listaexcel:
                                            if dicionario['Código Cliente'] == codigocliente and dicionario['Inscrição'] == linha['iptu']:
                                                dicionario.update({'Status': 'Verificar'})

                                    if site is not None:
                                        site.fecharsite()
                                else:
                                    dadosiptu = [codigocliente, linha['iptu'], '2022', 'Verificar (Extrair Manualmente)']
                                    listaexcel.append(dict(zip(listachaves, dadosiptu)))
                                    if site is not None:
                                        site.fecharsite()

                            else:
                                dadosiptu = [codigocliente, linha['iptu'], '2022', mensagemerro.text]
                                listaexcel.append(dict(zip(listachaves, dadosiptu)))
                                site.fecharsite()
        else:
            if os.path.isfile(caminhodestino) and salvardadospdf:
                listacodigo = []
                listatipopag = []
                listaarquivo = []
                df = aux.extrairtextopdf(caminhodestino)
                total_rows = df[df.columns[0]].count()
                for linha in range(total_rows):
                    listacodigo.append(codigocliente)
                    listatipopag.append('PARCELADO')
                    listaarquivo.append(caminhodestino)

                df.insert(loc=0, column='Codigo', value=listacodigo)
                df.insert(loc=4, column='TpoPagto', value=listatipopag)
                df.insert(loc=5, column='Arquivo', value=listaarquivo)

                bd.adicionardf('Codigos IPTUs', df, 7)
    # finally:
    print(codigocliente)
    if site is not None:
        site.fecharsite()

    if len(listaexcel) > 0:
        visual.mudartexto('labelstatus', 'Salvando lista...')
        aux.escreverlistaexcelog('Log_' + aux.acertardataatual() + '.xlsx', listaexcel)

    tempofim = time.time()

    hours, rem = divmod(tempofim-tempoinicio, 3600)
    minutes, seconds = divmod(rem, 60)

    visual.manipularradio(True)
    visual.acertaconfjanela(False)

    messagebox.msgbox(
        f'O tempo decorrido foi de: {"{:0>2}:{:0>2}:{:05.2f}".format(int(hours), int(minutes), int(seconds))}',
        messagebox.MB_OK, 'Tempo Decorrido')
