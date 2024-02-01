import pandas as pd
import os

def processar_texto(texto):
    linhas = texto.strip().split("\n")
    dados = []

    titulo_aba = None
    for linha in linhas:
        if linha.startswith("EIXO TEMÁTICO"):
            titulo_aba = "Eixo 2" #Aqui coloca o nome da aba da planilha
        else:
            partes = linha.split(" ", 1)
            if len(partes) > 1:
                topico, descricao = partes
                if "." in topico:
                    topico, subtópico = topico.split(".", 1)
                else:
                    subtópico = "*"
            else:
                topico = partes[0]
                subtópico = "*"
                descricao = ""

            if not subtópico.strip():
                subtópico = "0"

            dados.append([topico.strip(), subtópico.strip(), descricao.strip()])

    return dados, titulo_aba

# Texto a ser processado
texto = """
EIXO TEMÁTICO 1 - GESTÃO GOVERNAMENTAL E GOVERNANÇA PÚBLICA: ESTRATÉGIA, PESSOAS, PROJETOS E PROCESSOS
1 Planejamento e gestão estratégica: conceitos, princípios, etapas, níveis, métodos e ferramentas.
1.1 Balanced Scorecard (BSC).
1.2 Matriz SWOT.
1.3 Estabelecimento de objetivos e metas organizacionais.
1.4 Métodos de desdobramento de objetivos e metas e elaboração de planos de ação e mapas estratégicos.
1.5 Implementação de estratégias.
1.6 Análise de cenários.
1.7 Ferramentas de gestão.
1.8 Metodologias para medição de desempenho.
1.9 Indicadores de desempenho: conceito, formulação e análise.
1.10 Detalhamento da ferramenta de avaliação de desempenho: OKR.
2 Gestão de pessoas
2.1 Evolução e funções da gestão de pessoas.
2.2 Recrutamento e seleção.
2.3 Avaliação de desempenho e gestão do desempenho.
2.4 Valorização, sistemas de recompensas e responsabilização.
2.5 Indicadores de gestão de pessoas
2.6 Gestão por competências.
2.7 Gestão de redes organizacionais
2.8 Desenvolvimento gerencial.
2.9 Clima Organizacional.
2.10 Comportamento organizacional e cultura organizacional.
2.11 Grupos e equipes de trabalho.
2.12 Qualidade de vida no trabalho.
2.13 Flexibilidade organizacional e teletrabalho.
2.14 Gestão de Programas de Saúde.
2.15 Gestão da mudança: mudanças sociais, científicas, culturais e organizacionais.
2.16 Aprendizagem individual e aprendizagem organizacional.
2.17 Estratégias para gestão do autodesenvolvimento e gestão da aprendizagem organizacional.
2.18 Métodos, estratégias e tendências em treinamento, desenvolvimento e educação.
2.19 Diagnóstico de necessidades de treinamento.
2.20 Elaboração e gerenciamento de projetos e programas educacionais.
2.21 Teorias de aprendizagem e desenho/projeto instrucional.
2.22 Avaliação de treinamento.
2.23 Educação à distância.
2.24 Gestão do conhecimento.
2.25 Liderança; Estilos de liderança e situações de trabalho.
2.26 Teorias da motivação.
2.27 Negociação e gestão de conflitos.
2.28 Metodologias ágeis em gestão de pessoas.
2.29 Legislação de pessoal no serviço público.
2.30 Política Nacional de Desenvolvimento de Pessoas.
2.31 Tendências do futuro do serviço público.
3 Gestão de projetos.
3.1 Conceitos básicos.
3.2 Processos do PMBOK.
3.3 Gerenciamento da integração, do escopo, do tempo, de custos, da qualidade, de recursos humanos, de comunicações, de riscos, de aquisições, de partes interessadas.
3.4 Metodologias ágeis.
4 Gestão de processos.
4.1 Conceitos da abordagem por processos.
4.2 Técnicas de mapeamento, análise e melhoria de processos.
4.3 BPM.
4.4 Desenho de serviços públicos.
"""


dados_novo, aba_nova = processar_texto(texto)
arquivo_existente = "EditalVertical.xlsx" #Troque aqui o nome da planilha para comparar se existe

try:

    with pd.ExcelWriter(arquivo_existente, engine='openpyxl', mode='a') as writer:
        df_novo = pd.DataFrame(dados_novo, columns=["Tópico", "Subtópico", "Descrição"])
        df_novo.to_excel(writer, sheet_name=aba_nova, index=False)

except Exception as e:
    print("Erro ao adicionar dados ao arquivo Excel existente, criando um novo Excel.")

    df_novo = pd.DataFrame(dados_novo, columns=["Tópico", "Subtópico", "Descrição"])
    df_novo.to_excel("EditalVertical.xlsx", sheet_name= aba_nova, index=False) #Mude aqui o nome da planilha para salvar como desejar
