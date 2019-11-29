from Control import Control;
import time;
import sys;
import traceback;
from pprint import *;

#pega o caminho da pasta e o nome do xls de entrada
if len(sys.argv) > 1:
    caminho = sys.argv[1] + "\\";
    planilha = sys.argv[2];
else:
    caminho = "";
    planilha = "";

# inicializa os principais objetos
start = time.clock();
startDate = time.localtime();
try:    
    control = Control(plan_dados = planilha, path = caminho, time = startDate);
except:
    print("Erro de Execucao");
    print("Consulte o arquivo erro.txt");
    # cria o arquivo txt
    saidaResul = open(caminho + "erro.txt", "w");
    saidaResul.write(traceback.format_exc());
    sys.exit(1);
elapsed = time.clock();
elapsed = elapsed - start;

# exporta objetos do python para json se a opcao estiver habilitada na planilha
if control.isExpJsonHabilitada:
    control.exportaObjeto();

# libera a memoria
control = None;

print("Concluido - Tempo Total: " + str(int(float(elapsed))) + " segundos");