import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import os

def criar_relatorio():
    doc = docx.Document()

    for sec in doc.sections:
        sec.top_margin = Inches(0.8)
        sec.bottom_margin = Inches(0.8)
        sec.left_margin = Inches(0.8)
        sec.right_margin = Inches(0.8)

    # Cabecalho
    p_head = doc.add_paragraph()
    r_head = p_head.add_run(
        "CENTRO PAULA SOUZA - EAD TEC\n"
        "CURSO TECNICO EM DESENVOLVIMENTO DE SISTEMAS\n"
        "AGENDA 09 - TECNOLOGIAS DA INFORMACAO I (RECUPERACAO)"
    )
    r_head.bold = True
    r_head.font.size = Pt(10.5)
    r_head.font.color.rgb = RGBColor(70, 70, 70)
    p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_div = doc.add_paragraph()
    r_div = p_div.add_run("-" * 70)
    r_div.font.color.rgb = RGBColor(180, 180, 180)
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Titulo
    p_title = doc.add_paragraph()
    r_title = p_title.add_run("RELATORIO DE ENGENHARIA DE REQUISITOS E MODELAGEM AGIL")
    r_title.bold = True
    r_title.font.size = Pt(13)
    r_title.font.color.rgb = RGBColor(109, 40, 217)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_id = doc.add_paragraph()
    r_id = p_id.add_run("Aluno: Rodrigo Assis | Formato: Rodrigo_Ag9_TI_I_Recup")
    r_id.bold = True
    r_id.font.size = Pt(9.5)
    p_id.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()

    # 1. Requisitos Identificados
    doc.add_heading("1. Lista dos Requisitos Levantados (Empresa Doce Encanto)", level=1)
    
    p_rf = doc.add_paragraph()
    p_rf.add_run("Requisitos Funcionais (RF):\n").bold = True
    p_rf.add_run(
        "- RF01: O sistema deve permitir o cadastro e consulta de clientes.\n"
        "- RF02: O sistema deve registrar pedidos de encomendas com data, horario e itens solicitados.\n"
        "- RF03: O sistema deve disponibilizar uma agenda visual de entregas diárias e semanais.\n"
        "- RF04: O sistema deve disparar notificacao automatica de confirmacao de pedido ao cliente.\n"
        "- RF05: O sistema deve atualizar a baixa de ingredientes no estoque conforme os pedidos.\n"
        "- RF06: O sistema deve emitir relatorios financeiros de faturamento e fluxo de caixa."
    )

    p_rnf = doc.add_paragraph()
    p_rnf.add_run("Requisitos Nao Funcionais (RNF):\n").bold = True
    p_rnf.add_run(
        "- RNF01 (Usabilidade): Interface responsiva e simplificada para operacao rapida no balcao.\n"
        "- RNF02 (Desempenho): Tempo de resposta inferior a 2 segundos nas buscas e registros de pedidos.\n"
        "- RNF03 (Seguranca): Autenticacao por usuario e senha com backup diario na nuvem.\n"
        "- RNF04 (Compatibilidade): Acesso compativel com navegadores modernos e smartphones."
    )

    # 2. Caso de Uso Detalhado
    doc.add_heading("2. Caso de Uso Detalhado: Registrar Pedido de Encomenda", level=1)
    
    t_uc = doc.add_table(rows=1, cols=2)
    t_uc.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_uc.rows[0].cells[0].text = "Item de Especificacao"
    t_uc.rows[0].cells[1].text = "Detalhamento Tecnico"
    t_uc.rows[0].cells[0].paragraphs[0].runs[0].bold = True
    t_uc.rows[0].cells[1].paragraphs[0].runs[0].bold = True

    detalhes_uc = [
        ("Caso de Uso", "UC01 - Registrar Pedido de Encomenda"),
        ("Ator Principal", "Jessica (Proprietaria) ou Atendente de Balcao"),
        ("Objetivo de Uso", "Registrar com precisao o pedido do cliente, garantindo agendamento e baixa de insumos."),
        ("Pre-condicoes", "Usuario autenticado no sistema e catalogo de produtos cadastrado."),
        ("Fluxo Principal", 
         "1. O atendente acessa a tela de 'Novo Pedido';\n"
         "2. Seleciona ou cadastra o cliente;\n"
         "3. Adiciona os itens (bolos/doces) e quantidades desejadas;\n"
         "4. Define a data e hora limite para retirada/entrega;\n"
         "5. Informa a forma de pagamento e valor recebido;\n"
         "6. O sistema valida os dados, salva o pedido e agenda a producao;\n"
         "7. O sistema aciona o caso de uso 'Notificar Cliente' (include);\n"
         "8. O sistema confirma a gravacao com numero de protocolo."),
        ("Fluxos Alternativos / Excecoes", 
         "4a. Data sem disponibilidade na agenda: o sistema alerta sobrecarga e sugere horarios proximos;\n"
         "3a. Ingredientes insuficientes no estoque: o sistema emite aviso de estoque critico antes de confirmar.")
    ]

    for item, det in detalhes_uc:
        row = t_uc.add_row().cells
        row[0].text = item
        row[1].text = det

    doc.add_paragraph()

    # 3. Diagrama UML
    base_dir = os.path.dirname(os.path.abspath(__file__))
    img_diag = os.path.join(base_dir, "diagrama_caso_de_uso.png")
    if os.path.exists(img_diag):
        doc.add_heading("3. Diagrama de Casos de Uso (UML)", level=1)
        doc.add_picture(img_diag, width=Inches(5.8))
        p_cap = doc.add_paragraph("Figura 1: Diagrama de Casos de Uso do Sistema Doce Encanto.")
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.runs[0].font.size = Pt(8.5)
        p_cap.runs[0].italic = True

    # 4. Modelo Agil Proposto
    doc.add_heading("4. Proposta de Modelo Agil de Trabalho", level=1)
    doc.add_paragraph(
        "Para responder com agilidade as dores da proprietaria, adotou-se o Scrum com Sprints quinzenais:\n"
        "- Sprint 1: Modulo de Pedidos, Agenda e Notificacao (entrega de valor imediato para cessar atrasos);\n"
        "- Sprint 2: Gestao de Estoque de Ingredientes e Alertas de Ponto de Pedido;\n"
        "- Sprint 3: Modulo Financeiro, Fluxo de Caixa e Relatorios Gerenciais.\n"
        "A cliente Jessica atua diretamente como Product Owner nas sessoes quinzenais de Sprint Review."
    )

    # 5. Dialogo com IA
    img_ia = os.path.join(base_dir, "interacao_jessica.png")
    if os.path.exists(img_ia):
        doc.add_heading("5. Registro da Entrevista com a IA (Roleplay Jessica)", level=1)
        doc.add_picture(img_ia, width=Inches(5.8))
        p_cap2 = doc.add_paragraph("Figura 2: Dialogo simulado de levantamento de necessidades e validacao da proposta agil.")
        p_cap2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap2.runs[0].font.size = Pt(8.5)
        p_cap2.runs[0].italic = True

    out_file = os.path.join(base_dir, "Rodrigo_Ag9_TI_I_Recup.docx")
    doc.save(out_file)
    print("Documento salvo em:", out_file)

if __name__ == "__main__":
    criar_relatorio()
