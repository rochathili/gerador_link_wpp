import pandas as pd
from urllib.parse import quote
from pathlib import Path


# ==============================
# CONFIGURAÇÕES
# ==============================

DDD_PADRAO = "11"

COLUNAS_TELEFONE = ["Contato1", "Contato2", "Contato3"]

BASE_DIR = Path(__file__).resolve().parent

ARQUIVO_ENTRADA = BASE_DIR / "contatos.xlsx"

ARQUIVO_SAIDA = BASE_DIR / "links_wpp.xlsx"


# ==============================
# FUNÇÕES
# ==============================

def sanitizar_telefone(telefone):
    """
    Remove caracteres não numéricos e formata o telefone no padrão:

    55 + DDD + 9 + últimos 8 dígitos
    """

    if pd.isna(telefone):
        return None

    digitos = ''.join(filter(str.isdigit, str(telefone)))

    if len(digitos) < 8:
        return None

    ultimos_8 = digitos[-8:]

    return f"55{DDD_PADRAO}9{ultimos_8}"


def obter_telefone(row):
    """
    Tenta obter um telefone válido nas colunas:
    Contato1 → Contato2 → Contato3
    """

    for coluna in COLUNAS_TELEFONE:

        if coluna not in row:
            continue

        telefone_original = row[coluna]

        telefone_formatado = sanitizar_telefone(telefone_original)

        if telefone_formatado:
            return telefone_formatado, telefone_original

    return None, None


def gerar_link_whatsapp(telefone, mensagem):
    """
    Gera link do WhatsApp Web com a mensagem codificada para URL
    """

    mensagem = mensagem.replace("\n", "\r\n")
    mensagem_encoded = quote(mensagem, safe="")

    return f"https://web.whatsapp.com/send?phone={telefone}&text={mensagem_encoded}"


# ==============================
# MENSAGEM
# ==============================

MENSAGEM_TEMPLATE = """💜 *Parabéns!* 💜

{nome}, você foi *aprovado(a)* para o curso de *Desenvolvimento Web (Programação) NOITE* no *Instituto da Oportunidade Social* 💜

Para confirmar sua vaga, compareça para realizar sua *matrícula* entre os dias *11/03/2026 e 12/03/2026*, no período das *18h00 às 21h00* 💜

📍 *Local da matrícula:*

Instituto da Oportunidade Social  
Av. Gen. Ataliba Leonel, 245 - Santana  
São Paulo - SP, 02033-000

🗺️ Localização no mapa:  
https://maps.google.com/?q=Av.+Gen.+Ataliba+Leonel+245+Santana+Sao+Paulo

*Documentos necessários para matrícula:*

• Documento de identidade (RG)  
• CPF  
• Comprovante de residência  
• Comprovante de escolaridade  
• Comprovante de renda  

⚠️ Menores de idade deverão comparecer acompanhados por um responsável legal.

Aguardamos você para efetivar sua matrícula. 💜
"""


# ==============================
# PROCESSAMENTO
# ==============================

df = pd.read_excel(ARQUIVO_ENTRADA)

resultados = []

for _, row in df.iterrows():

    nome = row["Nome Completo"]

    telefone, telefone_original = obter_telefone(row)

    if not telefone:
        print(f"⚠️ {nome} - nenhum telefone válido")
        continue

    mensagem = MENSAGEM_TEMPLATE.format(nome=nome)

    link_wpp = gerar_link_whatsapp(telefone, mensagem)

    resultados.append({
        "Nome": nome,
        "Contato Original": telefone_original,
        "Telefone Formatado": telefone,
        "Link WhatsApp Web": link_wpp
    })

    print(f"✅ {nome}")
    print(f"📱 {telefone}")
    print(f"🔗 {link_wpp}\n")


# ==============================
# EXPORTAÇÃO
# ==============================

df_resultados = pd.DataFrame(resultados)

df_resultados.to_excel(ARQUIVO_SAIDA, index=False)

print("✔ Arquivo gerado com sucesso.")