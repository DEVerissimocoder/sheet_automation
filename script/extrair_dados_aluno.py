from bs4 import BeautifulSoup
import xlwt
texto_style = xlwt.XFStyle()
texto_style.num_format_str = '@'
import os

def extrair_dados_ficha_aluno(html):
    soup = BeautifulSoup(html, 'html.parser')
    dados = {}

    def encontrar_valor_por_label(label):
        td_label = soup.find('td', string=label)
        if td_label:
            td_valor = td_label.find_next_sibling('td')
            if td_valor:
                texto=td_valor.get_text(strip=True)
                texto= texto.replace('\xa0', '').strip()
            return texto
        return None

    dados['matricula'] = encontrar_valor_por_label('Matrícula:')
    dados['nome'] = encontrar_valor_por_label('Nome civil:')
    dados['data_nascimento'] = encontrar_valor_por_label('Data de nascimento:')
    dados['filiacao_1'] = encontrar_valor_por_label('Filiação 1:')
    dados['filiacao_2'] = encontrar_valor_por_label('Filiação 2:')  # pode ser None
    dados['email'] = encontrar_valor_por_label('Endereço eletrônico:')
    dados['telefone'] = encontrar_valor_por_label('Telefone:')
    return dados

# Pasta com os arquivos HTML
pasta_htmls = './alunos_html/1_A_noite'
arquivos_html = [f for f in os.listdir(pasta_htmls) if f.endswith('.html')]

# Criar planilha
workbook = xlwt.Workbook()
sheet = workbook.add_sheet('Aluno')

# Cabeçalhos fixos
cabecalhos = ['matricula', 'nome', 'data_nascimento', 'filiacao_1', 'filiacao_2', 'email', 'telefone']
for col, cab in enumerate(cabecalhos):
    sheet.write(0, col, cab)

# Processar cada arquivo HTML
for linha, nome_arquivo in enumerate(arquivos_html, start=1):
    caminho = os.path.join(pasta_htmls, nome_arquivo)
    with open(caminho, 'r', encoding='utf-8') as f:
        html = f.read()
    dados = extrair_dados_ficha_aluno(html)

    for col, cab in enumerate(cabecalhos):
        valor = dados.get(cab, '').strip()
        if cab == 'matricula':
            sheet.write(linha, col, valor, texto_style)  # força a matrícula como texto
        else:
            sheet.write(linha, col, valor)
# Salvar arquivo
workbook.save('1_A_N.xls')
print("Arquivo 'dados_aluno.xls' salvo com sucesso.")
