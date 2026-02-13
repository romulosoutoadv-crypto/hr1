!pip install python-docx num2words

import textwrap
from dataclasses import dataclass, field
from datetime import date, timedelta
from typing import List, Dict
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from num2words import num2words

# ==================== 1. MOTOR DE DECISÃO (PRESERVADO & EXPANDIDO) ====================
@dataclass
class Caso:
    # Dados do Cliente/Processo
    autor_nome: str
    num_processo: str
    comarca: str
    contrato_id: str
    valor_causa: float
    
    # Dados Fáticos para o Motor
    dt_contrato: date
    meses_pg: int
    dt_inad: date = None
    dt_acao: date = None
    
    # Gatilhos
    adit_25: bool = False      # Assinou aditivo de 2025?
    residence_club_reu: bool = False # Incluiu a Holding no polo passivo?
    gratuidade_deferida: bool = False # Juiz deu justiça gratuita?

class Motor_HRH_Decisao:
    @staticmethod
    def avaliar(c: Caso) -> Dict:
        teses = []
        
        # 1. Tese: Ilegitimidade Passiva (Baseado no Modelão e Caso Maria do Ceu)
        if c.residence_club_reu:
            teses.append("ILEGITIMIDADE_RESIDENCE")

        # 2. Tese: Impugnação à Justiça Gratuita (Baseado no Caso Maria do Ceu)
        if c.gratuidade_deferida and c.valor_causa > 40000:
             teses.append("IMPUGNACAO_AJG")

        # 3. Tese: Suppressio (Se pagou por muito tempo antes de processar)
        # Lógica: Se pagou mais de 24 meses e demorou a entrar com ação
        if c.meses_pg > 24:
            teses.append("SUPPRESSIO_FORTISSIMA")

        # 4. Tese: Caso Fortuito / Atraso (Padrão para contratos antigos)
        if c.dt_contrato.year <= 2021:
            teses.append("CASO_FORTUITO_PANDEMIA")
            
        # 5. Tese: Validade da Cláusula de Foro (Se não for em Fortaleza)
        if "Fortaleza" not in c.comarca:
            teses.append("INCOMPETENCIA_TERRITORIAL")

        return {"Teses": teses}

# ==================== 2. REPOSITÓRIO DE "DNA ESTILÍSTICO" (EXTRAÍDO DOS ARQUIVOS) ====================
# [span_0](start_span)[span_1](start_span)[span_2](start_span)[span_3](start_span)Textos retirados diretamente dos arquivos 'Modelão' [span_0](end_span)[span_1](end_span)[span_2](end_span)[span_3](end_span)

TEXTOS_PADRAO = {
    "HEADER": """
AO JUÍZO DA {comarca}
PROCESSO Nº {num_processo}
REQUERENTE: {autor_nome}
REQUERIDA: HRH FORTALEZA EMPREENDIMENTO HOTELEIRO S.A.

HRH FORTALEZA EMPREENDIMENTO HOTELEIRO S.A., atual denominação de VENTURE CAPITAL PARTICIPAÇÕES E INVESTIMENTOS S.A., pessoa jurídica de direito privado, vem, perante Vossa Excelência, apresentar a referida CONTESTAÇÃO.
    """,

    "INCOMPETENCIA_TERRITORIAL": """
    I. DA INCOMPETÊNCIA TERRITORIAL RELATIVA DO JUÍZO
    Preliminarmente, insta-se pelo declínio da competência, considerando que este juízo não é o competente para julgar a presente demanda, conforme estipulado contratualmente.
    As partes acordaram consensualmente que o foro competente seria o da comarca de Fortaleza/CE, renunciando a qualquer outro. O STJ já pacificou (Resp 1.675.012 SP) que a cláusula de eleição de foro só é inválida se provada a hipossuficiência, o que não ocorre no caso de multipropriedade de alto padrão.
    """,

    "ILEGITIMIDADE_RESIDENCE": """
    I. DA INCONTESTÁVEL ILEGITIMIDADE PASSIVA DO RESIDENCE CLUB S.A.
    Pugna-se pela ilegitimidade passiva da parte Requerida RESIDENCE CLUB S.A., uma vez que esta Sociedade não tem qualquer vínculo obrigacional com o Requerente.
    Trata-se de uma sociedade anônima fechada (Holding) cuja atividade principal é aluguel de imóveis próprios e participação em outras empresas, não se confundindo com a incorporadora HRH FORTALEZA. A inclusão desta no polo passivo configura erro grosseiro na formação da lide.
    """,

    "IMPUGNACAO_AJG": """
    I. DA IMPUGNAÇÃO À JUSTIÇA GRATUITA
    A parte Requerente demonstra um padrão de vida incompatível com a hipossuficiência alegada. O processo trata de fração imobiliária de alto padrão (Hard Rock Hotel). Quem assume tal compromisso e verte pagamentos expressivos não se enquadra no perfil que a legislação visa proteger.
    Requer-se, portanto, a revogação do benefício, sob pena de chancelar uma distorção do sistema legal.
    """,

    "SUPPRESSIO_FORTISSIMA": """
    II. DA SUPRESSIO E DO COMPORTAMENTO CONTRADITÓRIO (VENIRE CONTRA FACTUM PROPRIUM)
    Ocorre, no caso em tela, o fenômeno da 'suppressio'. A inércia da parte autora, que realizou pagamentos por longo período ({meses_pg} meses) sem qualquer ressalva administrativa, gerou na Requerida a legítima expectativa de continuidade do vínculo.
    Não pode agora, de forma contraditória, alegar inadimplemento pretérito que tacitamente perdoou ao continuar pagando as parcelas. Tal conduta viola a boa-fé objetiva (Art. 422 do CC).
    """,

    "CASO_FORTUITO_PANDEMIA": """
    III. DA OCORRÊNCIA DE FORTUITOS EXTERNOS (EXCLUDENTE DE RESPONSABILIDADE)
    O atraso alegado decorre de fatos necessários e inevitáveis: a pandemia de COVID-19 e a subsequente crise inflacionária na construção civil (INCC histórico).
    Tais eventos configuram fortuito externo, rompendo o nexo causal. A jurisprudência pátria e a própria Lei de Distratos autorizam a prorrogação de prazos em cenários de força maior, não havendo que se falar em culpa da incorporadora.
    """
}

# ==================== 3. PROTOCOLO DE GERAÇÃO (MIMETISMO) ====================
class Gerador_Contestacao:
    def __init__(self, output_path: str):
        self.doc = Document()
        self.output_path = output_path
        self._configurar_estilos()

    def _configurar_estilos(self):
        """Define o estilo 'HRH Standard' (Arial/Times, Justificado, 1.5)"""
        style = self.doc.styles['Normal']
        font = style.font
        font.name = 'Arial' # Fonte extraída do mimetismo
        font.size = Pt(11)
        
        par = style.paragraph_format
        par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        par.line_spacing = 1.5
        par.space_after = Pt(11)

    def adicionar_bloco(self, titulo_chave: str, texto_template: str, dados: Caso):
        """Injeta o texto formatando as variáveis dinamicamente"""
        # Formata valores monetários por extenso automaticamente
        texto_final = texto_template.format(
            comarca=dados.comarca.upper(),
            num_processo=dados.num_processo,
            autor_nome=dados.autor_nome.upper(),
            meses_pg=dados.meses_pg,
            valor_causa_extenso=f"R$ {dados.valor_causa:,.2f} ({num2words(dados.valor_causa, lang='pt_br', to='currency')})"
        )
        
        # Limpeza e inserção
        paragraphs = textwrap.dedent(texto_final).strip().split('\n')
        for p_text in paragraphs:
            if p_text.strip():
                p = self.doc.add_paragraph(p_text)
                # Se for título (começa com algarismo romano), negrito
                if any(p_text.strip().startswith(x) for x in ["I.", "II.", "III.", "AO JUÍZO"]):
                    p.runs[0].bold = True
                    if "AO JUÍZO" in p_text:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def salvar(self):
        self.doc.save(self.output_path)
        print(f"Documento '{self.output_path}' gerado com sucesso.")

# ==================== 4. EXECUÇÃO ====================
if __name__ == "__main__":
    # A. Configuração do Caso (Simulando dados do usuário)
    caso_usuario = Caso(
        autor_nome="Maria do Ceu Paiva de Oliveira",
        num_processo="0819194-55.2025.8.20.5106",
        comarca="Mossoró - RN", # Comarca diferente do foro (Fortaleza) -> Gatilho Incompetência
        contrato_id="H2-10309",
        valor_causa=69900.00,
        dt_contrato=date(2019, 11, 3), # Contrato antigo -> Gatilho Pandemia
        meses_pg=48, # Pagou muito tempo -> Gatilho Suppressio
        residence_club_reu=True, # Gatilho Ilegitimidade
        gratuidade_deferida=True # Gatilho Impugnação AJG
    )

    # B. Motor de Decisão
    decisao = Motor_HRH_Decisao.avaliar(caso_usuario)
    print(f"Teses Selecionadas pelo Motor: {decisao['Teses']}")

    # C. Geração do Documento
    gerador = Gerador_Contestacao("Contestacao_Automatica_HRH.docx")
    
    # 1. Cabeçalho
    gerador.adicionar_bloco("HEADER", TEXTOS_PADRAO["HEADER"], caso_usuario)
    
    # 2. Inserção das Teses Dinâmicas
    for tese_key in decisao["Teses"]:
        if tese_key in TEXTOS_PADRAO:
            gerador.adicionar_bloco(tese_key, TEXTOS_PADRAO[tese_key], caso_usuario)

    gerador.salvar()
