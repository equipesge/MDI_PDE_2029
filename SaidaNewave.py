from Sistema import Sistema;
from Problema import Problema;
from openpyxl import load_workbook;
from datetime import *;
from pyomo.environ import *;
from pyomo.environ import *;
from math import *;
import pyomo.opt;


class SaidaNewave:

    def __init__(self, sistema, problema, path, numerosSubsistemas, subsistemasNaoFicticios):
        # recebe como parametro o sistema em que estao as informacoes e o problema com o modelo
        self.sin = sistema;
        self.modelo = problema.modelo;
        self.caminho = path;
        self.numSubs = numerosSubsistemas;
        self.subsNFic = subsistemasNaoFicticios;

        # chama o metodo que cria a saida pequenas usinas antiga
        self.imprimePequenasUsinasAntigo();

        # metodos que criam as outras saidas
        self.imprimeExpansaoBinaria();
        self.imprimeExpansaoTransm();
        self.imprimeAgrint();
        self.imprimeExpansaoTermica();
        self.imprimeExpansaoHidreletrica();
        self.imprimePatamar();
        self.imprimeExpansaoTermica();
        self.imprimeSistema();

        return;

    def imprimePequenasUsinasAntigo(self):

        # declaracao de variaveis
        modelo = self.modelo;

        # cria o arquivo txt
        saidaResul = open(self.caminho + "newavePequenasUsinasAntigo.txt", "w"); 

        # imprime o cabecalho
        saidaResul.write("GERACAO DE PEQUENAS USINAS \n");
        saidaResul.write(" XXX \n");
        saidaResul.write("       XXXJAN. XXXFEV. XXXMAR. XXXABR. XXXMAI. XXXJUN. XXXJUL. XXXAGO. XXXSET. XXXOUT. XXXNOV. XXXDEZ. \n");

        # os subsistemas ficticios nao sao contados
        for isis in range (0, len(self.subsNFic)):

            # abreviacao para simplificar o entendimento
            subsis = self.modelo.sin.subsistemas[isis];
            saidaResul.write("  " + str(int(self.numSubs[isis])) + "\n");

            # zera o iano para escrever os mesmos anos em cada subsistema
            iano = 0;
            imes = 0;

            # escreve o ano no inicio da primeira linha
            saidaResul.write(str(self.sin.anoInicial) + "  ");

            for iper in (self.modelo.periodos):

                if (imes <= 11):

                    geracao = subsis.montanteRenovExPCH[iper]  \
                            + subsis.montanteRenovExEOL[iper]  \
                            + subsis.montanteRenovExUFV[iper]  \
                            + subsis.montanteRenovExBIO[iper]  \
                            + sum(modelo.capRenovCont[proj.nomeUsina,iper].value*proj.fatorCapacidade[iper % 12] for proj in subsis.listaProjUFV) \
                            + sum(modelo.capRenovCont[proj.nomeUsina,iper].value*proj.fatorCapacidade[iper % 12] for proj in subsis.listaProjBIO) \
                            + sum(modelo.capRenovCont[proj.nomeUsina,iper].value*proj.fatorCapacidade[iper % 12] for proj in subsis.listaProjEOF) \
                            + sum(modelo.capRenovCont[proj.nomeUsina,iper].value*proj.fatorCapacidade[iper % 12] for proj in subsis.listaProjPCH);

                    # adiciona a incerteza da serie de vento
                    if (self.sin.tipoCombHidroEol == "completa"):
                        geracao += (sum(modelo.capRenovCont[proj.nomeUsina,iper].value*proj.seriesEolicas[icond][iper % 12] for icond in self.modelo.condicoes for proj in subsis.listaProjEOL))/self.sin.numCondicoes;
                    elif (self.sin.tipoCombHidroEol == "intercalada"):
                        geracao += (sum(modelo.capRenovCont[proj.nomeUsina,iper].value*proj.seriesEolicas[icond % self.sin.numEol][iper % 12] for icond in self.modelo.condicoes for proj in subsis.listaProjEOL))/self.sin.numCondicoes;
                    else:
                        print("opcao de combinacao de series hidrologicas com eolicas nao marcada");        

                    # escreve a geracao
                    saidaResul.write('{0:>8s}'.format(str(round(geracao))+'.'));

                    # varia imes para controle da quebra de linha
                    imes += 1;
                    
                else:
                        
                    # pula a linha para escrever os dados referentes ao proximo ano
                    saidaResul.write("\n");

                    # escreve o ano no inicio de cada linha
                    iano += 1;
                    saidaResul.write(str(self.sin.anoInicial + iano) + "  ");

                    geracao = subsis.montanteRenovExPCH[iper]  \
                            + subsis.montanteRenovExEOL[iper]  \
                            + subsis.montanteRenovExUFV[iper]  \
                            + subsis.montanteRenovExBIO[iper]  \
                            + sum(modelo.capRenovCont[proj.nomeUsina,iper].value*proj.fatorCapacidade[iper % 12] for proj in subsis.listaProjUFV) \
                            + sum(modelo.capRenovCont[proj.nomeUsina,iper].value*proj.fatorCapacidade[iper % 12] for proj in subsis.listaProjBIO) \
                            + sum(modelo.capRenovCont[proj.nomeUsina,iper].value*proj.fatorCapacidade[iper % 12] for proj in subsis.listaProjEOF) \
                            + sum(modelo.capRenovCont[proj.nomeUsina,iper].value*proj.fatorCapacidade[iper % 12] for proj in subsis.listaProjPCH);

                    # adiciona a incerteza da serie de vento
                    if (self.sin.tipoCombHidroEol == "completa"):
                        geracao += (sum(modelo.capRenovCont[proj.nomeUsina,iper].value*proj.seriesEolicas[icond][iper % 12] for icond in self.modelo.condicoes for proj in subsis.listaProjEOL))/self.sin.numCondicoes;
                    elif (self.sin.tipoCombHidroEol == "intercalada"):
                        geracao += (sum(modelo.capRenovCont[proj.nomeUsina,iper].value*proj.seriesEolicas[icond % self.sin.numEol][iper % 12] for icond in self.modelo.condicoes for proj in subsis.listaProjEOL))/self.sin.numCondicoes;
                    else:
                        print("opcao de combinacao de series hidrologicas com eolicas nao marcada");    
                    
                    # escreve a geracao
                    saidaResul.write('{0:>8s}'.format(str(round(geracao))+'.'));

                    # varia imes para controle da quebra de linha
                    imes = 1;

            # pula a linha para escrever o proximo ano de fato
            saidaResul.write("\n");
            

        # fecha o arquivo txt
        saidaResul.close();

        return;

    def imprimeExpansaoBinaria(self):
        modelo = self.modelo;
        sin = self.sin;
        
        # abre o arquivo para a saidas
        saidaResul = open(self.caminho + "saidaExpansaoBinaria.txt", "w");
        
        # cabecalho
        saidaResul.write("HIDRO\n");
        
        # percorre todas as hidro
        for hidro in modelo.projUHENova:
            # pega o mes que a usina entrou
            mes_entrada = modelo.sin.numMeses # inicializa no final do horizonte
            for per in modelo.periodosTotal: # percorre todos os periodos
                if (modelo.investHidro[hidro,per] == 1): mes_entrada = per  # no periodo que entrou atribui o periodo
            
            # so imprime se a usina tiver entrado
            if mes_entrada < modelo.sin.numMeses:
                # imprime a entrada
                saidaResul.write(str(int(sin.listaGeralProjUHE[hidro].indexUsinaExterno)) + " no tempo " + str(int(mes_entrada+1)) + "\n");
        
        # fecha o arquivo
        saidaResul.close();
        
        return;

    def imprimeExpansaoTransm(self):
        modelo = self.modelo;
        sin = self.sin;
        
        # abre o arquivo para a saidas
        saidaResul = open(self.caminho + "saidaExpCapTransm.txt", "w");
        
        # cabecalho
        saidaResul.write("Mes;Sistema1;Sistema2;\n");

        # percorre os periodos
        for iper in (self.modelo.periodosTotal):
            # percorre os sistemas de envio
            for isis in modelo.subsistemas:
                # percorre os de recebimento
                for jsis in modelo.subsistemas:
                    # imprime a sequencia
                    saidaResul.write(str(iper+1) + ";" + str(self.numSubs[isis]) + ";" + str(self.numSubs[jsis]) + ";" + str(modelo.capExpInter[isis,jsis,iper].value) + "\n");
        
        saidaResul.close();
        
        return;

    def imprimeExpansaoTransmSistema(self):
        modelo = self.modelo;
        sin = self.sin;
        
        # abre o arquivo para a saidas
        saidaResul = open(self.caminho + "saidaExpTransmissaoSistema.txt", "w");
        saidaResul.write("\n");
        
        # percorre os sistemas de envio
        for isis in range(sin.nsis):
            # percorre os de recebimento
            for jsis in range(isis+1,sin.nsis):
                for ipat in range(0, sin.nPatamares):
                    # inicializa as linhas de impressao
                    linhasIJ = str(int(self.numSubs[isis])) + "  " + str(int(self.numSubs[jsis])) + "\n\n"
                    linhasJI = "";
                    # percorre os anos
                    for iano in range(sin.numAnos):
                        for imes in range(12):
                            # calcula o periodo
                            per = iano*12+imes;
                            
                            # concatena os valores nos dois sentidos nas duas strings;
                            linhasIJ += str(self.sin.subsistemas[isis].capExistente[jsis][ipat][per] + modelo.capExpInter[isis,jsis,per].value) + ";"
                            linhasJI += str(self.sin.subsistemas[jsis].capExistente[isis][ipat][per] + modelo.capExpInter[jsis,isis,per].value) + ";"
                        
                        # insere uma quebra de linha ao final do ano
                        linhasIJ += "\n";
                        linhasJI += "\n";
                        
                    # imprime os dois conjuntos de linha
                    saidaResul.write(linhasIJ + "\n");
                    saidaResul.write(linhasJI + "\n");
        
        saidaResul.close();
        
        return;

    def imprimePatamar(self):
        modelo = self.modelo;
        sin = self.sin;
        
        # abre o arquivo para a saidas
        saidaResul = open(self.caminho + "patamar.dat", "w");

        # imprime o cabecalho
        saidaResul.write(" NUMERO DE PATAMARES\n");
        saidaResul.write(" XX\n");
        saidaResul.write('{0:>3s}'.format(str(sin.nPatamares))+"\n");

        saidaResul.write("      DURACAO MENSAL DOS PATAMARES DE CARGA\n");
        saidaResul.write("ANO     JAN     FEV     MAR     ABR     MAI     JUN     JUL     AGO     SET     OUT     NOV     DEZ\n");
        saidaResul.write("XXXX  X.XXXX  X.XXXX  X.XXXX  X.XXXX  X.XXXX  X.XXXX  X.XXXX  X.XXXX  X.XXXX  X.XXXX  X.XXXX  X.XXXX\n");

        for iano in range(0,sin.numAnos):
            for ipat in range(0,sin.nPatamares):
                if ipat == 0:
                    linha = str(sin.anoInicial+iano);
                else:
                    linha = "    ";
                for imes in range(0,12):
                    linha += "  " + '{:6.4f}'.format(sin.duracaoPatamar[ipat][12*iano+imes]);
                saidaResul.write(linha + "\n");

        saidaResul.write("SUBSISTEMA\n");
        saidaResul.write(" XXX\n");
        saidaResul.write("                             CARGA(P.U.DEMANDA MED.)\n");
        saidaResul.write("        X.XXXX X.XXXX X.XXXX X.XXXX X.XXXX X.XXXX X.XXXX X.XXXX X.XXXX X.XXXX X.XXXX X.XXXX\n");

        for isis in range(0,len(self.subsNFic)): #nao imprime os ficticios
            saidaResul.write('{0:>4s}'.format(str(int(self.numSubs[isis])))+"\n");
            for iano in range(0,sin.numAnos):
                for ipat in range(0,sin.nPatamares):
                    if ipat == 0:
                        linha = "   " + str(sin.anoInicial+iano);
                    else:
                        linha = "       ";
                    for imes in range(0,12):
                        linha += " " + '{0:.4f}'.format(sin.subsistemas[isis].cargaPatamar[ipat][iano*12+imes]);
                    saidaResul.write(linha + "\n");

        saidaResul.write("9999\n");
        saidaResul.write(" SUBSISTEMA\n");
        saidaResul.write("   A ->B\n");
        saidaResul.write(" XXX XXX\n");
        saidaResul.write("                             INTERCAMBIO(P.U.INTERC.MEDIO)\n");
        saidaResul.write("        X.XXXX X.XXXX X.XXXX X.XXXX X.XXXX X.XXXX X.XXXX X.XXXX X.XXXX X.XXXX X.XXXX X.XXXX\n");
       
        # percorre os sistemas de envio
        for isis in range(sin.nsis):
            # percorre os de recebimento
            for jsis in range(isis+1,sin.nsis):
                # inicializa as linhas de impressao
                linhasIJ = " " + " "*(3-len(str(int(self.numSubs[isis])))) + str(int(self.numSubs[isis])) + " " + " "*(3-len(str(int(self.numSubs[jsis])))) + str(int(self.numSubs[jsis])) + "\n";
                linhasJI = " " + " "*(3-len(str(int(self.numSubs[jsis])))) + str(int(self.numSubs[jsis])) + " " + " "*(3-len(str(int(self.numSubs[isis])))) + str(int(self.numSubs[isis])) + "\n";

                for iano in range(0,sin.numAnos):
                    pIJ = [[0 for i in range(0, sin.nPatamares)] for j in range(12)];
                    pJI = [[0 for i in range(0, sin.nPatamares)] for j in range(12)];
                    for ipat in range(0,sin.nPatamares):
                        if ipat == 0:
                            linhasIJ += "   " + str(sin.anoInicial+iano);
                            linhasJI += "   " + str(sin.anoInicial+iano);
                        else:
                            linhasIJ += "       ";
                            linhasJI += "       ";
                        for imes in range(0,12):
                            if sum(sin.subsistemas[isis].capExistente[jsis][ipataux][12*iano+imes]*sin.duracaoPatamar[ipataux][12*iano+imes] for ipataux in modelo.patamares) + modelo.capExpInter[isis,jsis,12*iano+imes].value ==0 or sum(sin.subsistemas[jsis].capExistente[isis][ipataux][12*iano+imes]*sin.duracaoPatamar[ipataux][12*iano+imes] for ipataux in modelo.patamares) + modelo.capExpInter[jsis,isis,12*iano+imes].value ==0:
                                linhasIJ += " " + '{0:.4f}'.format(1);
                                linhasJI += " " + '{0:.4f}'.format(1);
                            else:
                                #correcao do patamar leve para manter a soma do produto igual a 1
                                if ipat == sin.nPatamares -1:
                                    pIJ[imes][ipat] = round((1-sum(pIJ[imes][ipataux]*sin.duracaoPatamar[ipataux][12*iano+imes] for ipataux in range(0,sin.nPatamares-1)))/sin.duracaoPatamar[sin.nPatamares-1][12*iano+imes],4);
                                    pJI[imes][ipat] = round((1-sum(pJI[imes][ipataux]*sin.duracaoPatamar[ipataux][12*iano+imes] for ipataux in range(0,sin.nPatamares-1)))/sin.duracaoPatamar[sin.nPatamares-1][12*iano+imes],4);
                                    linhasIJ += " " + '{0:.4f}'.format(pIJ[imes][ipat]);
                                    linhasJI += " " + '{0:.4f}'.format(pJI[imes][ipat]);
                                else:
                                    pIJ[imes][ipat] = round((sin.subsistemas[isis].capExistente[jsis][ipat][12*iano+imes] + modelo.capExpInter[isis,jsis,12*iano+imes].value)/(sum(sin.subsistemas[isis].capExistente[jsis][ipataux][12*iano+imes]*sin.duracaoPatamar[ipataux][12*iano+imes] for ipataux in modelo.patamares) + modelo.capExpInter[isis,jsis,12*iano+imes].value),4);
                                    pJI[imes][ipat] = round((sin.subsistemas[jsis].capExistente[isis][ipat][12*iano+imes] + modelo.capExpInter[jsis,isis,12*iano+imes].value)/(sum(sin.subsistemas[jsis].capExistente[isis][ipataux][12*iano+imes]*sin.duracaoPatamar[ipataux][12*iano+imes] for ipataux in modelo.patamares) + modelo.capExpInter[jsis,isis,12*iano+imes].value),4);
                                    linhasIJ += " " + '{0:.4f}'.format(pIJ[imes][ipat]);
                                    linhasJI += " " + '{0:.4f}'.format(pJI[imes][ipat]);
                        linhasIJ += "\n";
                        linhasJI += "\n";
                    # imprime os dois conjuntos de linha somente se o valor no final do horizonte for diferente de zero
                if sin.subsistemas[isis].capExistente[jsis][0][sin.numMeses-1] + modelo.capExpInter[isis,jsis,sin.numMeses-1].value > 0 or sin.subsistemas[jsis].capExistente[isis][0][sin.numMeses-1] + modelo.capExpInter[jsis,isis,sin.numMeses-1].value > 0:
                    saidaResul.write(linhasIJ);
                    saidaResul.write(linhasJI);

        saidaResul.write("9999\n");
        saidaResul.write(" SUB BLOCO\n");
        saidaResul.write(" XXX XXX\n");
        saidaResul.write("   ANO                   BLOCO DE USINAS NAO SIMULADAS (P.U. ENERGIA MEDIA)\n");
        saidaResul.write("   XXXX X.XXXX X.XXXX X.XXXX X.XXXX X.XXXX X.XXXX X.XXXX X.XXXX X.XXXX X.XXXX X.XXXX X.XXXX\n");

        # os subsistemas ficticios nao sao contados
        for isis in range(0,len(self.subsNFic)):

            # loop para cada tipo de usina
            for tipo in range (1, 6):

                # abreviacao para simplificar o entendimento
                subsis = self.modelo.sin.subsistemas[isis];

                # inicializa o bloco de impressao
                bloco = " " + " "*(3-len(str(int(isis+1)))) + str(int(isis+1)) + " " + " "*(3-len(str(int(tipo)))) + str(int(tipo)) + "\n";

                for iano in range(0,sin.numAnos):
                    for ipat in range(0,sin.nPatamares):
                        if ipat == 0:
                            bloco += "   " + str(sin.anoInicial+iano);
                        else:
                            bloco += "       ";
                        for imes in range(0,12):
                            if tipo == 1:
                                bloco += " " + '{0:.4f}'.format(sin.subsistemas[isis].fatorPatPCH[ipat][imes]);
                            if tipo == 2:
                                bloco += " " + '{0:.4f}'.format(sin.subsistemas[isis].fatorPatBIO[ipat][imes]);
                            if tipo == 3:
                                bloco += " " + '{0:.4f}'.format(sin.subsistemas[isis].fatorPatEOL[ipat][imes]);
                            if tipo == 4:
                                bloco += " " + '{0:.4f}'.format(sin.subsistemas[isis].fatorPatUFV[ipat][imes]);
                            if tipo == 5:
                                bloco += " " + '{0:.4f}'.format(sin.subsistemas[isis].fatorPatEOF[ipat][imes]);                            
                        bloco += "\n";
                # imprime o bloco
                saidaResul.write(bloco);
        saidaResul.write("9999");

        saidaResul.close();           

        return;

    def imprimeAgrint(self):
        modelo = self.modelo;
        sin = self.sin;
        
        # abre o arquivo para a saidas
        saidaResul = open(self.caminho + "saidaAgrint.txt", "w");
        
        # percorre os grupos de agrint
        for agrint_ind in modelo.conjAgrint:
            iano = 0;
            imes = 1;
            agrint = modelo.sin.agrints[agrint_ind];
            for iper in modelo.periodosTotal:
                linha = "";
                if imes <= 12: 

                    linha += '{0:>4s}'.format(str(agrint_ind + 1)); #numero agrint
                    linha += '{0:>4s}'.format(str(imes)); #mes inicial
                    linha += '{0:>5s}'.format(str(sin.anoInicial + iano)); #ano inicial
                    linha += '{0:>3s}'.format(str(imes)); #mes final
                    linha += '{0:>5s}'.format(str(sin.anoInicial + iano)); #ano final

                    #limites por patamar
                    for ipat in modelo.patamares:
                        if agrint.limites[ipat][iper] == 99999:
                            linha += '{0:>8s}'.format(str(round(agrint.limites[ipat][iper]))+".");
                        else:
                            linha += '{0:>8s}'.format(str(round(agrint.limites[ipat][iper] + sum(modelo.capExpInter[isis,jsis,iper].value for (isis,jsis) in agrint.fluxos)))+".");
                    
                    saidaResul.write(linha + "\n");

                    imes +=1;

                else:

                    iano += 1;
                    imes = 1;

                    linha += '{0:>4s}'.format(str(agrint_ind + 1)); #numero agrint
                    linha += '{0:>4s}'.format(str(imes)); #mes inicial
                    linha += '{0:>5s}'.format(str(sin.anoInicial + iano)); #ano inicial
                    linha += '{0:>3s}'.format(str(imes)); #mes final
                    linha += '{0:>5s}'.format(str(sin.anoInicial + iano)); #ano final

                    #limites por patamar
                    for ipat in modelo.patamares:
                        if agrint.limites[ipat][iper] == 99999:
                            linha += '{0:>8s}'.format(str(round(agrint.limites[ipat][iper]))+".");
                        else:
                            linha += '{0:>8s}'.format(str(round(agrint.limites[ipat][iper] + sum(modelo.capExpInter[isis,jsis,iper].value for (isis,jsis) in agrint.fluxos)))+".");
                    
                    saidaResul.write(linha + "\n");

                    imes += 1;
        
        saidaResul.close();
        
        return;

    def imprimeExpansaoTermica(self):
        modelo = self.modelo;
        sin = self.sin;
        
        # abre o arquivo para a saidas
        saidaResul = open(self.caminho + "saidaExpTermica.txt", "w");
        
        # cabecalho
        saidaResul.write("Termica;TempoEntrada;ExpAcumulado;ValorFim;\n");
       
        # pega a lista de nome de usinas ordenada pelo codigo
        listaProjTermOrdenadoCod = [usina for k,usina in self.sin.listaGeralProjTerm.items()];
        def getKey(u):
            return u.indexUsinaExterno;
        listaProjTermOrdenadoCod.sort(key=getKey);
        listaProjTermOrdenadoCod = [usina.nomeUsina for usina in listaProjTermOrdenadoCod];

        # percorre os projetos no modelo
        for term in listaProjTermOrdenadoCod:
            # pega o objeto do projeto
            projTerm = sin.listaGeralProjTerm[term];
            
            # cria um vetor com as linhas que vai escrever no arquivo
            linhasArquivo = [];
            nLinhas = 0;
            
            # percorre todos os periodos do segundo ao ultimo
            for per in range(1,sin.numMeses-1):
                # verifica se a capacidade alterou entre o mes corrente e o anterior
                if (modelo.capTermCont[term,per].value - modelo.capTermCont[term,per-1].value) > 0.01:
                    # insere a linha
                    linhasArquivo.append({'inicio': per, 'fim': -1, 'valor': modelo.capTermCont[term,per].value});
                    
                    # caso nao seja a primeira tem que finalizar a linha anterior
                    if nLinhas > 0:
                        linhasArquivo[nLinhas-1]['fim'] = per-1;
                    
                    # incrementa o numero de linhas
                    nLinhas += 1;
                    
            # depois de montar a estrutura de linhas imprime no arquivo cada uma
            for linha in linhasArquivo:
                # verifica se a linha tem fim para preencher
                fim = str(int(linha['fim']+1)) if linha['fim'] >= 0 else '';
                
                # imprime a linha            
                saidaResul.write(str(int(projTerm.indexUsinaExterno)) + ';' + str(int(linha['inicio']+1)) + ';' + str(linha['valor']) + ';' + fim + "\n");
                
        saidaResul.close();
        
        return;

    def imprimeExpansaoHidreletrica(self):

        # variaveis
        modelo = self.modelo;
        sin = self.sin;
        linha = "";

        # abre o arquivo para a saidas
        saidaResul = open(self.caminho + "saidaExpHidreletrica.txt", "w");
        
        # cabecalho
        saidaResul.write("NUM  NOME         POSTO   JUS REE  V.INIC U.EXIS MODIF INIC.HIST FIM HIS DESVIOS  H.IND\n");
        saidaResul.write("XXXX XXXXXXXXXXXX  XXXX  XXXX XXXX XXX.XX XXXX   XXXX     XXXX     XXXX XXXX XXXX XXXX");

        # pega a lista de nome de usinas ordenada pelo codigo
        listaProjUHEOrdenadoCod = [usina for k,usina in sin.listaGeralProjUHE.items()];
        def getKey(u):
            return u.indexUsinaExterno;
        listaProjUHEOrdenadoCod.sort(key=getKey);
        listaProjUHEOrdenadoCod = [usina.nomeUsina for usina in listaProjUHEOrdenadoCod];

        # percorre os projetos no modelo
        for uhe in listaProjUHEOrdenadoCod:
            # pega o objeto do projeto
            projUHE = sin.listaGeralProjUHE[uhe];

            linha += '{0:>4s}'.format(str(projUHE.indexUsinaExterno));
            linha += '{0:>4s}'.format(str(projUHE.nomeUsina)); 
            linha += '{0:>5s}'.format(str("posto"));  
            linha += '{0:>5s}'.format(str("jusante")); 
            linha += '{0:>5s}'.format(str(projUHE.sis_index));
            linha += '{0:>4s}'.format(str(0.00));
            linha += '{0:>4s}'.format(str("NE")); 
            linha += '{0:>5s}'.format(str(0)); 
            linha += '{0:>3s}'.format(str(1931)); 
            linha += '{0:>5s}'.format(str(int((projUHE.dataMinima/365)-2))); 

            # escreve no arquivo
            saidaResul.write(linha + "\n");

        # fecha o arquivo
        saidaResul.close();
        
        return;

    def imprimeExpansaoTermica(self):
        modelo = self.modelo;
        sin = self.sin;
        
        # abre os arquivos para a saidas
        saidaResul = open(self.caminho + "expt.dat", "w");
        saidaConft = open(self.caminho + "conft.dat", "w");
        
        # cabecalho
        saidaResul.write("NUS  TIPO  VALORNEW MI ANOI MF ANOF  COMENTARIO \n");
        saidaResul.write("XXXX XXXXX XXXXXXXX XX XXXX XX XXXX \n");

        saidaConft.write("  NUM         NOME   SSIS     EX CLASSE UTE IND\n");
        saidaConft.write(" XXXX XXXXXXXXXXXX   XXXX     XX   XXXX XXXX\n");

        # pega a lista de nome de usinas ordenada pelo codigo
        listaProjTermOrdenadoCod = [usina for k,usina in self.sin.listaGeralProjTerm.items()];
        def getKey(u):
            return u.indexUsinaExterno;
        listaProjTermOrdenadoCod.sort(key=getKey);
        listaProjTermOrdenadoCod = [usina.nomeUsina for usina in listaProjTermOrdenadoCod];
	
        # percorre os projetos no modelo
        for term in listaProjTermOrdenadoCod:
            # pega o objeto do projeto
            projTerm = sin.listaGeralProjTerm[term];
            
            # cria um vetor com as linhas que vai escrever no arquivo
            linhasArquivo = [];
            nLinhas = 0;
            flagDivideUsina = False;
            
            # percorre todos os periodos do segundo ao ultimo
            for per in range(1,sin.numMeses):
                # verifica se a capacidade alterou entre o mes corrente e o anterior
                if (modelo.capTermCont[term,per].value*projTerm.potUsina - modelo.capTermCont[term,per-1].value * projTerm.potUsina) > 0.01:
                    # insere a linha
                    # divide a usina caso ultrapasse 9999 MW
                    if (modelo.capTermCont[term,per].value*projTerm.potUsina) > 9999 and flagDivideUsina == False:
                        flagDivideUsina = True;
                        linhasArquivo.append({'cod': projTerm.indexUsinaExterno, 'inicio': per, 'fim': -1, 'valor': 9999});
                        nLinhas += 1;
                        linhasArquivo.append({'cod': projTerm.indexUsinaExterno + 1, 'inicio': per, 'fim': -1, 'valor': modelo.capTermCont[term,per].value * projTerm.potUsina / projTerm.fdisp - 9999});
                    elif modelo.capTermCont[term,per].value*projTerm.potUsina > 9999 and flagDivideUsina == True:
                        linhasArquivo.append({'cod': projTerm.indexUsinaExterno + 1,'inicio': per, 'fim': -1, 'valor': modelo.capTermCont[term,per].value * projTerm.potUsina / projTerm.fdisp - 9999});
                    else:
                        linhasArquivo.append({'cod': projTerm.indexUsinaExterno, 'inicio': per, 'fim': -1, 'valor': modelo.capTermCont[term,per].value * projTerm.potUsina / projTerm.fdisp});
                    
                    # caso nao seja a primeira tem que finalizar a linha anterior
                    if nLinhas > 0 and linhasArquivo[nLinhas-1]['valor'] != 9999:
                        linhasArquivo[nLinhas-1]['fim'] = per-1;
                    
                    # incrementa o numero de linhas
                    nLinhas += 1;
            
            # inicia flag que imprime nome da usina somente na primeira linha
            flagImpressaoNomeUsina = False;

            # depois de montar a estrutura de linhas imprime no arquivo cada uma
            for linha in linhasArquivo:
                # imprime as linhas POTEF
                saidaResul.write('{:>4}'.format(str(int(linha['cod']))) + ' POTEF ' + '{:8.2f}'.format(linha['valor']));
                # se é dezembro o resto da divisao da 0
                if (linha['inicio']+1)%12 == 0:
                    saidaResul.write('{:>3}'.format(str(12)));
                else:
                    saidaResul.write('{:>3}'.format(str((linha['inicio']+1)%12)));
                saidaResul.write('{:>5}'.format(str(sin.anoInicial + (linha['inicio']//12))));
                # verifica se a linha tem fim para preencher
                if linha['fim'] >= 0:
                    # se é dezembro o resto da divisao da 0
                    if (linha['fim']+1)%12 == 0:
                        saidaResul.write('{:>3}'.format(str(12)));
                    else:
                        saidaResul.write('{:>3}'.format(str((linha['fim']+1)%12)));
                    saidaResul.write('{:>5}'.format(str(sin.anoInicial + (linha['fim']//12))));
                else:
                    saidaResul.write(8 * ' ');
                if flagImpressaoNomeUsina == False:
                    saidaResul.write('  ' + projTerm.nomeUsina + "\n");
                    flagImpressaoNomeUsina = True;
                else:
                    saidaResul.write("\n");

                # imprime as linhas GTMIN
                # diferencia as usinas com inflex sazonal das demais
                if projTerm.inflexSazonal == False:
                    # se a inflex for 0 nao precisa imprimir
                    if projTerm.inflexContinua[0] > 0:
                        saidaResul.write('{:>4}'.format(str(int(linha['cod']))) + ' GTMIN ' + '{:8.2f}'.format(linha['valor'] * projTerm.fdisp * projTerm.inflexContinua[0]));
                        # se é dezembro o resto da divisao da 0
                        if (linha['inicio']+1)%12 == 0:
                            saidaResul.write('{:>3}'.format(str(12)));
                        else:
                            saidaResul.write('{:>3}'.format(str((linha['inicio']+1)%12)));
                        saidaResul.write('{:>5}'.format(str(sin.anoInicial + (linha['inicio']//12))));
                        # verifica se a linha tem fim para preencher
                        if linha['fim'] >= 0:
                            # se é dezembro o resto da divisao da 0
                            if (linha['fim']+1)%12 == 0:
                                saidaResul.write('{:>3}'.format(str(12)));
                            else:
                                saidaResul.write('{:>3}'.format(str((linha['fim']+1)%12)));
                            saidaResul.write('{:>5}'.format(str(sin.anoInicial + (linha['fim']//12))));
                        else:
                            saidaResul.write(8*' ');
                        saidaResul.write("\n");
                else:
                    # se o fim for -1 será igual ao ultimo periodo
                    if linha['fim'] == -1:
                        fim = sin.numMeses;
                    else:
                        fim = linha['fim'] + 1;
                    for iper in range(linha['inicio'],fim):
                        # se a inflex for 0 nao precisa imprimir
                        if projTerm.inflexContinua[iper%12] > 0:
                            saidaResul.write('{:>4}'.format(str(int(linha['cod']))) + ' GTMIN ' + '{:8.2f}'.format(modelo.capTermCont[term,iper].value * projTerm.potUsina * projTerm.fdisp * projTerm.inflexContinua[iper%12]));
                            # se é dezembro o resto da divisao da 0
                            if (iper+1)%12 == 0:
                                saidaResul.write('{:>3}'.format(str(12)));
                            else:
                                saidaResul.write('{:>3}'.format(str(iper%12+1)));
                            saidaResul.write('{:>5}'.format(str(sin.anoInicial + iper//12)));
                            # fim é o mesmo mes
                            if (iper+1)%12 == 0:
                                saidaResul.write('{:>3}'.format(str(12)));
                            else:
                                saidaResul.write('{:>3}'.format(str(iper%12+1)));
                            saidaResul.write('{:>5}'.format(str(sin.anoInicial + iper//12)));
                            saidaResul.write("\n");

            saidaConft.write('{:>5}'.format(int(projTerm.indexUsinaExterno)) + ' ' + '{:12.12}'.format(projTerm.nomeUsina) + '{:>7}'.format(str(int(self.numSubs[int(projTerm.sis_index)-1]))));
            if flagImpressaoNomeUsina == True:
                saidaConft.write('{:>7}'.format('NE') + '{:>7}'.format(int(projTerm.indexUsinaExterno)) + '{:>5}'.format('s') + "\n");
            else:
                saidaConft.write('{:>7}'.format('NC') + '{:>7}'.format(int(projTerm.indexUsinaExterno)) + '{:>5}'.format('s') + "\n");
            if flagDivideUsina == True:
                saidaConft.write('{:>5}'.format(int(projTerm.indexUsinaExterno+1)) + ' ' + '{:12.12}'.format(projTerm.nomeUsina) + '{:>7}'.format(str(int(self.numSubs[int(projTerm.sis_index)-1]))));
                saidaConft.write('{:>7}'.format('NE') + '{:>7}'.format(int(projTerm.indexUsinaExterno+1)) + '{:>5}'.format('s') + "\n");
            
        saidaResul.close();
        saidaConft.close();
        
        return;

    def imprimeSistema(self):
        modelo = self.modelo;
        sin = self.sin;
        
        # abre o arquivo para a saidas
        saidaResul = open(self.caminho + "sistema.dat", "w");

        # imprime o cabecalho
        saidaResul.write(" LIMITES DE INTERCAMBIO\n");
        saidaResul.write(" A   B   A->B    B->A\n");
        saidaResul.write(" XXX XXX XJAN. XXXFEV. XXXMAR. XXXABR. XXXMAI. XXXJUN. XXXJUL. XXXAGO. XXXSET. XXXOUT. XXXNOV. XXXDEZ.\n");
        
        # percorre os sistemas de envio
        for isis in range(sin.nsis):
            # percorre os de recebimento
            for jsis in range(isis+1,sin.nsis):
                # inicializa as linhas de impressao
                linhasIJ = " " + " "*(3-len(str(int(self.numSubs[isis])))) + str(int(self.numSubs[isis])) + " " + " "*(3-len(str(int(self.numSubs[jsis])))) + str(int(self.numSubs[jsis])) + "\n";
                linhasJI = "";
                
                # zera o iano para escrever os mesmos anos em cada subsistema
                iano = 0;
                imes = 0;

                # escreve o ano no inicio da primeira linha
                linhasIJ += str(sin.anoInicial) + "  ";
                linhasJI += str(sin.anoInicial) + "  ";

                for iper in (modelo.periodos):

                    if (imes <= 11):

                        capIJ = sum(sin.subsistemas[isis].capExistente[jsis][ipat][iper]*sin.duracaoPatamar[ipat][iper] for ipat in modelo.patamares) + modelo.capExpInter[isis,jsis,iper].value;
                        capJI = sum(sin.subsistemas[jsis].capExistente[isis][ipat][iper]*sin.duracaoPatamar[ipat][iper] for ipat in modelo.patamares) + modelo.capExpInter[jsis,isis,iper].value;

                        linhasIJ += '{0:>8s}'.format(str(round(capIJ))+'.');
                        linhasJI += '{0:>8s}'.format(str(round(capJI))+'.');

                        if imes == 11:
                            # insere uma quebra de linha ao final do ano
                            linhasIJ += "\n";
                            linhasJI += "\n";

                        imes += 1;

                    else:

                        # escreve o ano no inicio de cada linha
                        iano += 1;
                        imes = 1;

                        linhasIJ += str(sin.anoInicial + iano) + "  ";
                        linhasJI += str(sin.anoInicial + iano) + "  ";

                        capIJ = sum(sin.subsistemas[isis].capExistente[jsis][ipat][iper]*sin.duracaoPatamar[ipat][iper] for ipat in modelo.patamares) + modelo.capExpInter[isis,jsis,iper].value;
                        capJI = sum(sin.subsistemas[jsis].capExistente[isis][ipat][iper]*sin.duracaoPatamar[ipat][iper] for ipat in modelo.patamares) + modelo.capExpInter[jsis,isis,iper].value;

                        linhasIJ += '{0:>8s}'.format(str(round(capIJ))+'.');
                        linhasJI += '{0:>8s}'.format(str(round(capJI))+'.');
                    
                # imprime os dois conjuntos de linha somente se o valor no final do horizonte for diferente de zero
                if sin.subsistemas[isis].capExistente[jsis][0][sin.numMeses-1] + modelo.capExpInter[isis,jsis,sin.numMeses-1].value or sin.subsistemas[jsis].capExistente[isis][0][sin.numMeses-1] + modelo.capExpInter[jsis,isis,sin.numMeses-1].value > 0:
                    saidaResul.write(linhasIJ + "\n");
                    saidaResul.write(linhasJI);
        
        saidaResul.write(" 999\n");

        # imprime o cabecalho
        saidaResul.write(" MERCADO DE ENERGIA TOTAL\n");
        saidaResul.write(" XXX\n");
        saidaResul.write("       XXXJAN. XXXFEV. XXXMAR. XXXABR. XXXMAI. XXXJUN. XXXJUL. XXXAGO. XXXSET. XXXOUT. XXXNOV. XXXDEZ.\n");

        for isis in range(0,len(self.subsNFic)): #nao imprime os ficticios
            subs = self.sin.subsistemas[isis]
            saidaResul.write('{0:>4s}'.format(str(int(self.numSubs[isis])))+"\n");
            for iano in range(0,sin.numAnos+1):
                if iano != sin.numAnos:
                    linha = str(sin.anoInicial+iano) + "  ";
                    for imes in range(12):
                        #verifica se tem reversivel neste subsistema
                        cap_rev = 0;
                        for rev in modelo.projReversivel:
                            proj = self.sin.listaGeralProjReversivel[rev]; # pega o projeto                        
                            # so imprime se for do submercado em questao
                            if proj.sis_index == (isis+1):
                                cap_rev += modelo.capReversivel[rev, iano*12+imes].value*(1/modelo.rendReversivel[rev]-1);
                        demanda_adic = cap_rev*sum(sin.duracaoPatamar[ipat][iano*12+imes] for ipat in modelo.sin.naoBombReversivel);
                        linha += '{:>7}'.format(round(subs.demandaMedia[iano*12+imes] + demanda_adic)) + ".";
                else:
                    linha = linha.replace(str(sin.anoInicial + sin.numAnos -1), "POS ");
                saidaResul.write(linha + "\n");

        saidaResul.write(" 999\n");

        escreveTipo = ["  01   BLOCO PCH", "  02   BLOCO PCT", "  03   BLOCO UEE", "  04   BLOCO UFV", "  05   BLOCO EOF"];    

        # imprime o cabecalho
        saidaResul.write(" GERACAO DE USINAS NAO SIMULADAS \n");
        saidaResul.write(" XXX XXX  XXXXXXXXXXXXXXXXXXXX \n");
        saidaResul.write("       XXXJAN. XXXFEV. XXXMAR. XXXABR. XXXMAI. XXXJUN. XXXJUL. XXXAGO. XXXSET. XXXOUT. XXXNOV. XXXDEZ. \n");

        # subsistemas ficticios nao sao contados
        for isis in range(0,len(self.subsNFic)):

            # loop para cada tipo de usina
            for tipo in range (0, 5):

                # abreviacao para simplificar o entendimento
                subsis = self.modelo.sin.subsistemas[isis];

                # so escreve se for diferente de zero
                if (tipo == 0) and (subsis.montanteRenovExPCH[sin.numMeses-1] + sum(modelo.capRenovCont[proj.nomeUsina,sin.numMeses-1].value for proj in subsis.listaProjPCH) > 0):

                    saidaResul.write("   " + str(int(self.numSubs[isis])) + escreveTipo[tipo] + "\n");

                    # zera o iano para escrever os mesmos anos em cada subsistema
                    iano = 0;
                    imes = 0;

                    # escreve o ano no inicio da primeira linha
                    saidaResul.write(str(self.sin.anoInicial) + "  ");

                    for iper in (self.modelo.periodos):

                        if (imes <= 11):
                            geracao = subsis.montanteRenovExPCH[iper] \
                                + sum(modelo.capRenovCont[proj.nomeUsina,iper].value*proj.fatorCapacidade[iper % 12] for proj in subsis.listaProjPCH);

                            saidaResul.write('{0:>8s}'.format(str(round(geracao))+'.'));

                            # varia imes para controle da quebra de linha
                            imes += 1;
                            
                        else:
                                
                            # pula a linha para escrever os dados referentes ao proximo ano
                            saidaResul.write("\n");

                            # escreve o ano no inicio de cada linha
                            iano += 1;
                            saidaResul.write(str(self.sin.anoInicial + iano) + "  ");

                            geracao = subsis.montanteRenovExPCH[iper] \
                                + sum(modelo.capRenovCont[proj.nomeUsina,iper].value*proj.fatorCapacidade[iper % 12] for proj in subsis.listaProjPCH);

                            saidaResul.write('{0:>8s}'.format(str(round(geracao))+'.'));

                            # varia imes para controle da quebra de linha
                            imes = 1;

                    # pula a linha para escrever o proximo ano de fato
                    saidaResul.write("\n");
                if (tipo == 1) and (subsis.montanteRenovExBIO[sin.numMeses-1] + sum(modelo.capRenovCont[proj.nomeUsina,sin.numMeses-1].value for proj in subsis.listaProjBIO) > 0):

                    saidaResul.write("   " + str(int(self.numSubs[isis])) + escreveTipo[tipo] + "\n");

                    # zera o iano para escrever os mesmos anos em cada subsistema
                    iano = 0;
                    imes = 0;

                    # escreve o ano no inicio da primeira linha
                    saidaResul.write(str(self.sin.anoInicial) + "  ");

                    for iper in (self.modelo.periodos):

                        if (imes <= 11):
                            geracao = subsis.montanteRenovExBIO[iper] \
                                + sum(modelo.capRenovCont[proj.nomeUsina,iper].value*proj.fatorCapacidade[iper % 12] for proj in subsis.listaProjBIO);

                            saidaResul.write('{0:>8s}'.format(str(round(geracao))+'.'));

                            # varia imes para controle da quebra de linha
                            imes += 1;
                            
                        else:
                                
                            # pula a linha para escrever os dados referentes ao proximo ano
                            saidaResul.write("\n");

                            # escreve o ano no inicio de cada linha
                            iano += 1;
                            saidaResul.write(str(self.sin.anoInicial + iano) + "  ");

                            geracao = subsis.montanteRenovExBIO[iper] \
                                + sum(modelo.capRenovCont[proj.nomeUsina,iper].value*proj.fatorCapacidade[iper % 12] for proj in subsis.listaProjBIO);

                            saidaResul.write('{0:>8s}'.format(str(round(geracao))+'.'));

                            # varia imes para controle da quebra de linha
                            imes = 1;

                    # pula a linha para escrever o proximo ano de fato
                    saidaResul.write("\n");

                if (tipo == 2) and (subsis.montanteRenovExEOL[sin.numMeses-1] + sum(modelo.capRenovCont[proj.nomeUsina,sin.numMeses-1].value for proj in subsis.listaProjEOL) > 0):

                    saidaResul.write("   " + str(int(self.numSubs[isis])) + escreveTipo[tipo] + "\n");

                    # zera o iano para escrever os mesmos anos em cada subsistema
                    iano = 0;
                    imes = 0;

                    # escreve o ano no inicio da primeira linha
                    saidaResul.write(str(self.sin.anoInicial) + "  ");

                    for iper in (self.modelo.periodos):

                        if (imes <= 11):
                            geracao = subsis.montanteRenovExEOL[iper];

                            if (self.sin.tipoCombHidroEol == "completa"):
                                geracao += (sum(modelo.capRenovCont[proj.nomeUsina,iper].value*proj.seriesEolicas[icond][iper % 12] for icond in self.modelo.condicoes for proj in subsis.listaProjEOL))/self.sin.numCondicoes;
                            elif (self.sin.tipoCombHidroEol == "intercalada"):
                                geracao += (sum(modelo.capRenovCont[proj.nomeUsina,iper].value*proj.seriesEolicas[icond % self.sin.numEol][iper % 12] for icond in self.modelo.condicoes for proj in subsis.listaProjEOL))/self.sin.numCondicoes;
                            else:
                                print("opcao de combinacao de series hidrologicas com eolicas nao marcada");    
                            
                            saidaResul.write('{0:>8s}'.format(str(round(geracao))+'.'));

                            # varia imes para controle da quebra de linha
                            imes += 1;
                            
                        else:
                                
                            # pula a linha para escrever os dados referentes ao proximo ano
                            saidaResul.write("\n");

                            # escreve o ano no inicio de cada linha
                            iano += 1;
                            saidaResul.write(str(self.sin.anoInicial + iano) + "  ");


                            geracao = subsis.montanteRenovExEOL[iper];
                        
                            if (self.sin.tipoCombHidroEol == "completa"):
                                geracao += (sum(modelo.capRenovCont[proj.nomeUsina,iper].value*proj.seriesEolicas[icond][iper % 12] for icond in self.modelo.condicoes for proj in subsis.listaProjEOL))/self.sin.numCondicoes;
                            elif (self.sin.tipoCombHidroEol == "intercalada"):
                                geracao += (sum(modelo.capRenovCont[proj.nomeUsina,iper].value*proj.seriesEolicas[icond % self.sin.numEol][iper % 12] for icond in self.modelo.condicoes for proj in subsis.listaProjEOL))/self.sin.numCondicoes;
                            else:
                                print("opcao de combinacao de series hidrologicas com eolicas nao marcada");    

                            saidaResul.write('{0:>8s}'.format(str(round(geracao))+'.'));

                            # varia imes para controle da quebra de linha
                            imes = 1;

                    # pula a linha para escrever o proximo ano de fato
                    saidaResul.write("\n");

                if (tipo == 3) and (subsis.montanteRenovExUFV[sin.numMeses-1] + sum(modelo.capRenovCont[proj.nomeUsina,sin.numMeses-1].value for proj in subsis.listaProjUFV) > 0):

                    saidaResul.write("   " + str(int(self.numSubs[isis])) + escreveTipo[tipo] + "\n");

                    # zera o iano para escrever os mesmos anos em cada subsistema
                    iano = 0;
                    imes = 0;

                    # escreve o ano no inicio da primeira linha
                    saidaResul.write(str(self.sin.anoInicial) + "  ");

                    for iper in (self.modelo.periodos):

                        if (imes <= 11):
                            geracao = subsis.montanteRenovExUFV[iper] \
                                + sum(modelo.capRenovCont[proj.nomeUsina,iper].value*proj.fatorCapacidade[iper % 12] for proj in subsis.listaProjUFV);

                            saidaResul.write('{0:>8s}'.format(str(round(geracao))+'.'));

                            # varia imes para controle da quebra de linha
                            imes += 1;
                            
                        else:
                                
                            # pula a linha para escrever os dados referentes ao proximo ano
                            saidaResul.write("\n");

                            # escreve o ano no inicio de cada linha
                            iano += 1;
                            saidaResul.write(str(self.sin.anoInicial + iano) + "  ");

                            geracao = subsis.montanteRenovExUFV[iper] \
                                + sum(modelo.capRenovCont[proj.nomeUsina,iper].value*proj.fatorCapacidade[iper % 12] for proj in subsis.listaProjUFV);

                            saidaResul.write('{0:>8s}'.format(str(round(geracao))+'.'));

                            # varia imes para controle da quebra de linha
                            imes = 1;

                    # pula a linha para escrever o proximo ano de fato
                    saidaResul.write("\n");

                if (tipo == 4) and sum(modelo.capRenovCont[proj.nomeUsina,sin.numMeses-1].value for proj in subsis.listaProjEOF) > 0:

                    saidaResul.write("   " + str(int(self.numSubs[isis])) + escreveTipo[tipo] + "\n");

                    # zera o iano para escrever os mesmos anos em cada subsistema
                    iano = 0;
                    imes = 0;

                    # escreve o ano no inicio da primeira linha
                    saidaResul.write(str(self.sin.anoInicial) + "  ");

                    for iper in (self.modelo.periodos):

                        if (imes <= 11):
                            geracao = sum(modelo.capRenovCont[proj.nomeUsina,iper].value*proj.fatorCapacidade[iper % 12] for proj in subsis.listaProjEOF);

                            saidaResul.write('{0:>8s}'.format(str(round(geracao))+'.'));

                            # varia imes para controle da quebra de linha
                            imes += 1;
                            
                        else:
                                
                            # pula a linha para escrever os dados referentes ao proximo ano
                            saidaResul.write("\n");

                            # escreve o ano no inicio de cada linha
                            iano += 1;
                            saidaResul.write(str(self.sin.anoInicial + iano) + "  ");

                            geracao = sum(modelo.capRenovCont[proj.nomeUsina,iper].value*proj.fatorCapacidade[iper % 12] for proj in subsis.listaProjEOF);

                            saidaResul.write('{0:>8s}'.format(str(round(geracao))+'.'));

                            # varia imes para controle da quebra de linha
                            imes = 1;

                    # pula a linha para escrever o proximo ano de fato
                    saidaResul.write("\n");

        # fecha o arquivo txt
        saidaResul.close();
        
        return;