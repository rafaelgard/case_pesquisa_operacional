import os
import numpy as np
import pandas as pd
import gurobipy as grb
from pyomo.environ import *
from pyomo.environ import Var
from pyomo.contrib.appsi.solvers.highs import Highs

class case:
    def __init__(self, CONFIGURACAO):
        self.nome_arquivo = CONFIGURACAO['nome_arquivo']
        self.tempo_exec = CONFIGURACAO['tempo_limite_de_execucao']
        self.solver = CONFIGURACAO['solver']
        self.quantidade_maxima_de_ondas = CONFIGURACAO['quantidade_maxima_de_ondas']
        self.modelo = None
        self.pre_processamento()

    def pre_processamento(self):
        print('Carregando arquivos...')
        self.carrega_arquivos()
        print('Arquivos carregados...')

        caminho_arquivo = 'output/solucao_final.xlsx'

        # Verifica se o arquivo existe antes de tentar apagá-lo
        if os.path.exists(caminho_arquivo):
            os.remove(caminho_arquivo)
            print("Solução antiga excluida...")
               
        print('Gerando modelo...')
        self.gera_modelo()
        print('Modelo gerado...')

    
    def executar_modelo(self):

        print(f'Otimizando o modelo com o solver {self.solver}...')
        if self.solver == 'gurobi':


            self.modelo.optimize()

            if self.modelo.SolCount > 0:

                print('O modelo encontrou solução factível.')

                if self.modelo.status == grb.GRB.OPTIMAL:
                    print("Solução ótima encontrada!")
                    print("Imprimindo resultados:")

                    self.salva_solucao()
            
            else:
                print('O modelo não encontrou uma solução factível.')
                exit()

        elif self.solver == 'highs':
            
            optimizer = Highs()

            optimizer.highs_options["time_limit"] = self.tempo_exec

            '''Otimiza o modelo'''
            try:
                results = optimizer.solve(self.modelo)

                print('O modelo encontrou solução factível.')
                self.salva_solucao()

            except:
                print('O modelo não encontrou uma solução factível.')
                exit()

        else: 
            print('Para rodar o modelo é necessário escolher um solver!\n')
            print('Opções disponíveis: "gurobi" e "highs')
            exit()

    def carrega_arquivos(self):
        '''Esta função carrega os arquivos das abas da planilha'''

        #Verifica se o arquivo de imput existe na pasta imput
        if os.path.exists('imput/'+self.nome_arquivo):

            self.df = pd.read_excel('imput/'+self.nome_arquivo)

            if not all(self.df.columns == ['Caixa Id', 'Item', 'Peças']):
                print('A aba está com as colunas diferente do padrão!')
                print('O correto deveria ser: "Caixa Id",	"Item",	"Peças"')
                print(f'Corrija o arquivo {self.nome_arquivo} e rode o modelo novamente.')
                exit()

        else:
            print(f'O arquivo "{self.nome_arquivo}" não foi encontrado na pasta imput! Verifique se o nome do arquivo está correto e sua existencia na pasta imput.')
            exit()

    def salva_solucao(self):
        '''Salva a solução do modelo em um arquivo excel'''

        if self.solver == 'gurobi':
            if self.modelo.status == grb.GRB.OPTIMAL:

                # Coletando resultados
                resultados_caixas = []
                for i in self.caixas:
                    for j in self.ondas:
                        if self.modelo.getVarByName(f"x[{i},{j}]").x > 0.5:
                            resultados_caixas.append([i, j])
                
                resultados_itens = []
                for k in self.itens:
                    for j in self.ondas:
                        if self.modelo.getVarByName(f"z[{k},{j}]").x > 0.5:
                            resultados_itens.append([k, j])
                                
            else:
                print("Nenhuma solução ótima encontrada.")

        elif self.solver == 'highs':

            resultados_caixas = []
            for i in self.caixas:
                for j in self.ondas:
                        valor = self.modelo.x[i,j].value
                        if valor > 0.5:
                            resultados_caixas.append([i, j])

            resultados_itens = []
            for j in self.ondas:
                for k in self.itens:
                    valor = self.modelo.z[k,j].value
                    if valor > 0.5:
                        resultados_itens.append([k, j])

        df_caixas = pd.DataFrame(resultados_caixas, columns=["Caixa Id", "Onda"])
        df_itens = pd.DataFrame(resultados_itens, columns=["Item", "Onda"])
        
        # Salva a solução
        nome_arquivo_solucao="solucao_final.xlsx"

        with pd.ExcelWriter('output/' + nome_arquivo_solucao) as writer:
            df_caixas.to_excel(writer, sheet_name='Caixas por Onda', index=False)
            df_itens.to_excel(writer, sheet_name='Itens por Onda', index=False)
        
        print(f"Solução salva em {'output/'+ nome_arquivo_solucao}")
                    
    
    def gera_modelo(self):

        # Parâmetros
        self.caixas = self.df['Caixa Id'].unique()
        self.itens = self.df['Item'].unique()
        self.ondas = range(1, self.quantidade_maxima_de_ondas)
        capacidade = 2000  # Capacidade máxima por onda

        if self.solver == 'gurobi':

            print('Iniciando modelo...')
    
            # Inicializando o modelo
            model = grb.Model('Modelo de Organização de Ondas de Produção')

            #===================================================================Variáveis de Decisão===================================================================
            x = model.addVars(self.caixas, self.ondas, vtype=grb.GRB.BINARY, lb=0, ub=1, name="x")  # Caixa i na onda j
            y = model.addVars(self.ondas, vtype=grb.GRB.BINARY, lb=0, ub=1, name="y")  # Onda j está ativa
            z = model.addVars(self.itens, self.ondas, vtype=grb.GRB.BINARY, lb=0, ub=1, name="z")  # Item k na onda j

            # Função Objetivo: Minimizar o número de ondas por item
            model.setObjective(
                grb.quicksum(z[k, j] for k in self.itens for j in self.ondas) / len(self.itens),
                grb.GRB.MINIMIZE
            )

            #===================================================================RESTRIÇÕES===================================================================

            # exige que cada caixa 𝑖 esteja atrelada a uma onda 𝑗.
            for i in self.caixas:
                model.addConstr(grb.quicksum(x[i, j] for j in self.ondas) == 1)

            # exige que a caixa 𝑖 só pode ser atrelada à onda 𝑗 se a onda estiver ativa.
            for i in self.caixas:
                for j in self.ondas:
                    model.addConstr(x[i, j] <= y[j])

            # faz a atribuição do item 𝑘, da lista de itens da caixa 𝑖, à onda 𝑗, se a caixa estiver atrelada à onda. 𝐾𝑖 define a lista de itens da caixa 𝑖
            for index, row in self.df.iterrows():
                item = row['Item']
                caixa = row['Caixa Id']
                for j in self.ondas:
                    model.addConstr(z[item, j] >= x[caixa, j])

            # exige que a soma total de peças 𝑝𝑖 das caixas atreladas à onda 𝑗 seja menor ou igual à capacidade 𝑐 (2000 peças).
            for j in self.ondas:
                model.addConstr(grb.quicksum(x[i, j] * self.df[self.df['Caixa Id'] == i]['Peças'].sum() for i in self.caixas) <= capacidade * y[j])
            
            # limite de tempo em segundos
            model.Params.timeLimit = self.tempo_exec

            model.update()

            self.modelo = model
  
        elif self.solver == 'highs':

            print('Iniciando modelo...')

            model = ConcreteModel()

            #===================================================================Variáveis de Decisão===================================================================
            model.x = Var(self.caixas, self.ondas, within=Binary, name='x')

            model.y = Var(self.ondas, within=Binary, name='y')

            model.z = Var(self.itens, self.ondas, within=Binary, name='z')

            '''Define a função objetivo'''
            model.obj = Objective(expr=(quicksum(model.z[k, j] for k in self.itens for j in self.ondas) / len(self.itens)),
                                sense = minimize
                                )
            
            #===================================================================RESTRIÇÕES===================================================================

            # exige que cada caixa 𝑖 esteja atrelada a uma onda 𝑗.
            model.R1 = ConstraintList()
            for i in self.caixas:
                model.R1.add(expr=(quicksum(model.x[i, j] for j in self.ondas)==1))

            # exige que a caixa 𝑖 só pode ser atrelada à onda 𝑗 se a onda estiver ativa.
            model.R2 = ConstraintList()
            for i in self.caixas:
                for j in self.ondas:
                    model.R2.add(expr=(model.x[i, j] <= model.y[j]))

            # faz a atribuição do item 𝑘, da lista de itens da caixa 𝑖, à onda 𝑗, se a caixa estiver atrelada à onda. 𝐾𝑖 define a lista de itens da caixa 𝑖
            model.R3 = ConstraintList()
            for index, row in self.df.iterrows():
                item = row['Item']
                caixa = row['Caixa Id']
                for j in self.ondas:
                    model.R3.add(expr=(model.z[item, j] >= model.x[caixa, j]))

            # exige que a soma total de peças 𝑝𝑖 das caixas atreladas à onda 𝑗 seja menor ou igual à capacidade 𝑐 (2000 peças).
            model.R4 = ConstraintList()
            for j in self.ondas:
                model.R4.add(expr=(quicksum(model.x[i, j] * self.df[self.df['Caixa Id'] == i]['Peças'].sum() for i in self.caixas) <= capacidade * model.y[j]))

            self.modelo = model

        else: 
            print('Para rodar o modelo é necessário escolher um solver!\n')
            print('Opções disponíveis: "gurobi" e "highs')
            exit()

#==============================================================================
nome_arquivo = 'Teste Pesquisa Operacional - Dados.xlsx'

CONFIGURACAO = {
                'nome_arquivo':nome_arquivo,
                'tempo_limite_de_execucao':3600,
                'solver':'highs',
                'quantidade_maxima_de_ondas': 30
                }

aloc = case(CONFIGURACAO)
aloc.executar_modelo()
#==============================================================================