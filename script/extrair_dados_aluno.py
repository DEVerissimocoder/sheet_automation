from bs4 import BeautifulSoup
import xlwt

def extrair_dados_ficha_aluno(html):
    soup = BeautifulSoup(html, 'html.parser')
    dados = {}

    def encontrar_valor_por_label(label):
        td_label = soup.find('td', string=label)
        if td_label:
            td_valor = td_label.find_next_sibling('td')
            return td_valor.get_text(strip=True) if td_valor else None
        return None

    dados['matricula'] = encontrar_valor_por_label('Matrícula:')
    dados['nome'] = encontrar_valor_por_label('Nome civil:')
    dados['data_nascimento'] = encontrar_valor_por_label('Data de nascimento:')
    dados['raca_cor'] = encontrar_valor_por_label('Raça/cor:')
    dados['filiacao_1'] = encontrar_valor_por_label('Filiação 1:')
    dados['responsavel'] = encontrar_valor_por_label('Responsável:')  # pode ser None
    dados['email'] = encontrar_valor_por_label('Endereço eletrônico:')

    return dados

# Abrir HTML
with open('./alunos_html/1_A/adriano.html', 'r', encoding='utf-8') as f:
    html = f.read()

# Extrair dados
dados = extrair_dados_ficha_aluno(html)

# Criar planilha
workbook = xlwt.Workbook()
sheet = workbook.add_sheet('Aluno')

# Escrever cabeçalhos
for col, chave in enumerate(dados.keys()):
    sheet.write(0, col, chave)

# Escrever dados na linha 1
for col, valor in enumerate(dados.values()):
    sheet.write(1, col, valor)

# Salvar arquivo
workbook.save('dados_aluno.xls')
print("Arquivo 'dados_aluno.xls' salvo com sucesso.")
