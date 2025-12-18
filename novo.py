import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import hashlib
import io
from datetime import datetime
import re
from collections import defaultdict
import xml.dom.minidom as minidom
import os
import zipfile
import time
from datetime import datetime
import pytz # Importar pytz



@st.cache_data
def parse_xte(file):
    file.seek(0)
    content = file.read().decode('iso-8859-1')
    tree = ET.ElementTree(ET.fromstring(content))
    root = tree.getroot()
    ns = {'ans': 'http://www.ans.gov.br/padroes/tiss/schemas'}
    all_data = []
    
    # Coleta as informações do cabecalho uma vez
    cabecalho_info = {}
    cabecalho = root.find('.//ans:cabecalho', namespaces=ns)
    if cabecalho is not None:
        identificacao = cabecalho.find('ans:identificacaoTransacao', namespaces=ns)
        if identificacao is not None:
            cabecalho_info['tipoTransacao'] = identificacao.findtext('ans:tipoTransacao', default='', namespaces=ns)
            cabecalho_info['numeroLote'] = identificacao.findtext('ans:numeroLote', default='', namespaces=ns)
            cabecalho_info['competenciaLote'] = identificacao.findtext('ans:competenciaLote', default='', namespaces=ns)
            cabecalho_info['dataRegistroTransacao'] = identificacao.findtext('ans:dataRegistroTransacao', default='', namespaces=ns)
            cabecalho_info['horaRegistroTransacao'] = identificacao.findtext('ans:horaRegistroTransacao', default='', namespaces=ns)
        cabecalho_info['registroANS'] = cabecalho.findtext('ans:registroANS', default='', namespaces=ns)
        cabecalho_info['versaoPadrao'] = cabecalho.findtext('ans:versaoPadrao', default='', namespaces=ns)

    for guia in root.findall(".//ans:guiaMonitoramento", namespaces=ns):
        guia_data = {}
        guia_data.update(cabecalho_info)

        # Loop principal para ler todas as tags como texto
        for elem in guia.iter():
            tag_full = elem.tag.split('}')[-1]
            if 'data' in tag_full.lower() and elem.text:
                try:
                    date_obj = datetime.strptime(elem.text, '%Y-%m-%d')
                    guia_data[tag_full] = date_obj.strftime('%d/%m/%Y')
                except ValueError:
                    guia_data[tag_full] = elem.text
            else:
                guia_data[tag_full] = elem.text if elem.text else None

        procedimentos = guia.findall(".//ans:procedimentos", namespaces=ns)
        if procedimentos:
            for proc in procedimentos:
                proc_data = guia_data.copy()
                # Extração específica dos procedimentos
                proc_data['codigoProcedimento'] = (proc.findtext('ans:identProcedimento/ans:Procedimento/ans:codigoProcedimento', namespaces=ns) or '').strip()
                proc_data['grupoProcedimento'] = (proc.findtext('ans:identProcedimento/ans:Procedimento/ans:grupoProcedimento', namespaces=ns) or '').strip()
                proc_data['valorInformado'] = (proc.findtext('ans:valorInformado', namespaces=ns) or '').strip()
                proc_data['valorPagoProc'] = (proc.findtext('ans:valorPagoProc', namespaces=ns) or '').strip()
                campos_procedimento = ['quantidadeInformada', 'quantidadePaga', 'valorPagoFornecedor', 'valorCoParticipacao', 'unidadeMedida']
                for campo in campos_procedimento:
                    proc_data[campo] = (proc.findtext(f'ans:{campo}', namespaces=ns) or '').strip()
                proc_data['codigoTabela'] = (proc.findtext('ans:identProcedimento/ans:codigoTabela', namespaces=ns) or '').strip()
                proc_data['registroANSOperadoraIntermediaria'] = (proc.findtext('ans:registroANSOperadoraIntermediaria', namespaces=ns) or '').strip()
                proc_data['tipoAtendimentoOperadoraIntermediaria'] = (proc.findtext('ans:tipoAtendimentoOperadoraIntermediaria', namespaces=ns) or '').strip()
                all_data.append(proc_data)
        else:
            all_data.append(guia_data)

    df = pd.DataFrame(all_data)
    df['Nome da Origem'] = file.name

    date_columns = [col for col in df.columns if 'data' in col.lower()]
    for col in date_columns:
        try:
            df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce').dt.strftime('%d/%m/%Y')
        except Exception:
            pass

    # Calcular idade
    if 'dataRealizacao' in df.columns and 'dataNascimento' in df.columns:
        def calcular_idade(row):
            try:
                data_realizacao = datetime.strptime(row['dataRealizacao'], '%d/%m/%Y')
                data_nascimento = datetime.strptime(row['dataNascimento'], '%d/%m/%Y')
                return (data_realizacao - data_nascimento).days // 365
            except Exception:
                return None
        df['Idade_na_Realização'] = df.apply(calcular_idade, axis=1)

    # --- BLOCO DE REMOÇÃO DE ZEROS À ESQUERDA FOI REMOVIDO DAQUI ---

    # --- Padronização e ordenação das colunas (sem alterações) ---
    df.rename(columns={
        'valorInformado': 'valorInformado_proc',
        'valorPagoFornecedor': 'valorPagoFornecedor_proc',
        'dataRegistroTransacao': 'dataRegistroTransacao_cabecalho',
        'horaRegistroTransacao': 'horaRegistroTransacao_cabecalho',
        'registroANS': 'registroANS_cabecalho',
        'versaoPadrao': 'versaoPadrao_cabecalho'
    }, inplace=True)

    colunas_finais = [
        'Nome da Origem', 'tipoRegistro', 'versaoTISSPrestador', 'formaEnvio', 'tipoTransacao',
        'numeroLote', 'competenciaLote', 'dataRegistroTransacao_cabecalho', 'horaRegistroTransacao_cabecalho',
        'registroANS_cabecalho', 'versaoPadrao_cabecalho', 'CNES', 'identificadorExecutante',
        'codigoCNPJ_CPF', 'municipioExecutante', 'registroANSOperadoraIntermediaria',
        'tipoAtendimentoOperadoraIntermediaria', 'numeroCartaoNacionalSaude', 'cpfBeneficiario',
        'sexo', 'dataNascimento', 'municipioResidencia', 'numeroRegistroPlano',
        'tipoEventoAtencao', 'origemEventoAtencao', 'numeroGuia_prestador', 'numeroGuia_operadora',
        'identificacaoReembolso', 'formaRemuneracao', 'valorRemuneracao', 'guiaSolicitacaoInternacao',
        'dataSolicitacao', 'numeroGuiaSPSADTPrincipal', 'dataAutorizacao', 'dataRealizacao',
        'dataFimPeriodo','dataInicialFaturamento', 'dataProtocoloCobranca', 'dataPagamento', 'dataProcessamentoGuia',
        'tipoConsulta', 'cboExecutante', 'indicacaoRecemNato', 'indicacaoAcidente',
        'caraterAtendimento', 'tipoInternacao', 'regimeInternacao', 'tipoAtendimento',
        'regimeAtendimento', 'tipoFaturamento', 'diariasAcompanhante', 'diariasUTI', 'motivoSaida',
        'valorTotalInformado', 'valorProcessado', 'valorTotalPagoProcedimentos', 'valorTotalDiarias',
        'valorTotalTaxas', 'valorTotalMateriais', 'valorTotalOPME', 'valorTotalMedicamentos',
        'valorGlosaGuia', 'valorPagoGuia', 'valorPagoFornecedores', 'valorTotalTabelaPropria',
        'valorTotalCoParticipacao', 'declaracaoNascido', 'declaracaoObito', 'codigoTabela',
        'grupoProcedimento', 'codigoProcedimento', 'quantidadeInformada','valorInformado', 'valorInformado_proc',
        'valorPagoFornecedor','quantidadePaga', 'unidadeMedida','valorCoParticipacao', 'valorPagoProc', 'valorPagoFornecedor_proc',
        'Idade_na_Realização', 'diagnosticoCID'
    ]

    for col in colunas_finais:
        if col not in df.columns:
            df[col] = None
    
    df = df[colunas_finais]

    return df, content, tree
    

def remove_duplicate_columns(df):
    df = df.loc[:, ~df.columns.duplicated()]
    df = df.dropna(axis=1, how='all')
    return df


def gerar_xte_do_excel(excel_file):
    print("--- DEBUG: Gerando XTE com lote por Minuto e Segundo (versão completa) ---")
    ns = "http://www.ans.gov.br/padroes/tiss/schemas"

    # --- Setup de Data/Hora e Leitura do Arquivo ---
    fuso_horario_servidor = pytz.utc
    fuso_horario_desejado = pytz.timezone("America/Sao_Paulo")
    agora_no_fuso_desejado = datetime.now(fuso_horario_servidor).astimezone(fuso_horario_desejado)
    data_atual = agora_no_fuso_desejado.strftime("%Y-%m-%d")
    hora_atual = agora_no_fuso_desejado.strftime("%H:%M:%S")

    # AJUSTE FINAL: Trocando Hora (%H) por Minuto (%M) na composição do lote.
    minuto_e_segundos_atuais = agora_no_fuso_desejado.strftime("%M%S")

    if hasattr(excel_file, 'name') and excel_file.name.endswith('.csv'):
        df = pd.read_csv(excel_file, dtype=str, sep=';')
    else:
        df = pd.read_excel(excel_file, dtype=str)

    # --- Função Auxiliar 'sub' ---
    def sub(parent, tag, value, is_date=False):
        if pd.isna(value):
            return
        text = str(value).strip()
        if is_date and text:
            original_text = text
            text = original_text
            for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
                try:
                    text = datetime.strptime(original_text, fmt).strftime("%Y-%m-%d")
                    break
                except ValueError:
                    continue
        if text:
            ET.SubElement(parent, f"ans:{tag}").text = text

    def extrair_texto(elemento):
        textos = []
        if elemento.text: textos.append(elemento.text.strip())
        for filho in elemento:
            textos.extend(extrair_texto(filho))
            if filho.tail: textos.append(filho.tail.strip())
        return textos

    arquivos_gerados = {}
    if "Nome da Origem" not in df.columns:
        raise ValueError("A coluna 'Nome da Origem' é obrigatória no Excel.")

    # --- Início da Geração do XML ---
    for nome_arquivo, df_origem in df.groupby("Nome da Origem"):
        if df_origem.empty: continue

        agrupado = df_origem.groupby(
            ["numeroGuia_prestador", "numeroGuia_operadora", "identificacaoReembolso"], dropna=False
        )
        
        root = ET.Element("ans:mensagemEnvioANS", attrib={
            "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance", "xmlns:xsd": "http://www.w3.org/2001/XMLSchema",
            "xsi:schemaLocation": f"{ns} {ns}/tissMonitoramentoV1_04_01.xsd", "xmlns:ans": ns
        })
        
        linha_cabecalho = df_origem.iloc[0]
        
        # --- Bloco do Cabeçalho ---
        cabecalho = ET.SubElement(root, "ans:cabecalho")
        identificacaoTransacao = ET.SubElement(cabecalho, "ans:identificacaoTransacao")
        sub(identificacaoTransacao, "tipoTransacao", "MONITORAMENTO")
        
        # AJUSTE FINAL: Geração do numeroLote com Minuto e Segundo
        competencia = linha_cabecalho.get("competenciaLote", "")
        if competencia and len(competencia) == 6 and competencia.isdigit():
            numero_lote_final = f"{competencia}{minuto_e_segundos_atuais}"
        else:
            ano_e_mes_atuais = agora_no_fuso_desejado.strftime("%Y%m")
            numero_lote_final = f"{ano_e_mes_atuais}{minuto_e_segundos_atuais}"

        sub(identificacaoTransacao, "numeroLote", numero_lote_final)
        sub(identificacaoTransacao, "competenciaLote", linha_cabecalho.get("competenciaLote"))
        sub(identificacaoTransacao, "dataRegistroTransacao", data_atual)
        sub(identificacaoTransacao, "horaRegistroTransacao", hora_atual)
        sub(cabecalho, "registroANS", linha_cabecalho.get("registroANS_cabecalho"))
        sub(cabecalho, "versaoPadrao", linha_cabecalho.get("versaoPadrao_cabecalho", "1.04.01"))

        mensagem = ET.SubElement(root, "ans:Mensagem")
        op_ans = ET.SubElement(mensagem, "ans:operadoraParaANS")

        # --- Loop Principal para cada Guia ---
        for _, grupo_guia_key in agrupado:
            linha_guia = grupo_guia_key.iloc[0]
            guia = ET.SubElement(op_ans, "ans:guiaMonitoramento")

            # --- Mapeamento Estruturado da Guia (Nível da Guia) ---
            sub(guia, "tipoRegistro", linha_guia.get("tipoRegistro"))
            sub(guia, "versaoTISSPrestador", linha_guia.get("versaoTISSPrestador"))
            sub(guia, "formaEnvio", linha_guia.get("formaEnvio"))
            
            dadosContratadoExecutante_el = ET.SubElement(guia, "ans:dadosContratadoExecutante")
            sub(dadosContratadoExecutante_el, "CNES", linha_guia.get("CNES"))
            sub(dadosContratadoExecutante_el, "identificadorExecutante", linha_guia.get("identificadorExecutante"))
            sub(dadosContratadoExecutante_el, "codigoCNPJ_CPF", linha_guia.get("codigoCNPJ_CPF"))
            sub(dadosContratadoExecutante_el, "municipioExecutante", linha_guia.get("municipioExecutante"))

            sub(guia, "registroANSOperadoraIntermediaria", linha_guia.get("registroANSOperadoraIntermediaria"))
            sub(guia, "tipoAtendimentoOperadoraIntermediaria", linha_guia.get("tipoAtendimentoOperadoraIntermediaria"))

            dadosBeneficiario_el = ET.SubElement(guia, "ans:dadosBeneficiario")
            identBeneficiario_el = ET.SubElement(dadosBeneficiario_el, "ans:identBeneficiario")
            sub(identBeneficiario_el, "numeroCartaoNacionalSaude", linha_guia.get("numeroCartaoNacionalSaude"))
            sub(identBeneficiario_el, "cpfBeneficiario", linha_guia.get("cpfBeneficiario"))
            sub(identBeneficiario_el, "sexo", linha_guia.get("sexo"))
            sub(identBeneficiario_el, "dataNascimento", linha_guia.get("dataNascimento"), is_date=True)
            sub(identBeneficiario_el, "municipioResidencia", linha_guia.get("municipioResidencia"))
            sub(dadosBeneficiario_el, "numeroRegistroPlano", linha_guia.get("numeroRegistroPlano"))

            sub(guia, "tipoEventoAtencao", linha_guia.get("tipoEventoAtencao"))
            sub(guia, "origemEventoAtencao", linha_guia.get("origemEventoAtencao"))
            sub(guia, "numeroGuia_prestador", linha_guia.get("numeroGuia_prestador"))
            sub(guia, "numeroGuia_operadora", linha_guia.get("numeroGuia_operadora"))
            
            origem_evento = linha_guia.get("origemEventoAtencao")
            valor_reembolso = ""
            if origem_evento in ['1', '2', '3']:
                valor_reembolso = "00000000000000000000"
            else:
                valor_reembolso = linha_guia.get("identificacaoReembolso")
            sub(guia, "identificacaoReembolso", valor_reembolso)
            
            if pd.notna(linha_guia.get("formaRemuneracao")) or pd.notna(linha_guia.get("valorRemuneracao")):
                formasRemuneracao_el = ET.SubElement(guia, "ans:formasRemuneracao")
                sub(formasRemuneracao_el, "formaRemuneracao", linha_guia.get("formaRemuneracao"))
                sub(formasRemuneracao_el, "valorRemuneracao", linha_guia.get("valorRemuneracao"))
            
            sub(guia, "guiaSolicitacaoInternacao", linha_guia.get("guiaSolicitacaoInternacao"))
            sub(guia, "dataSolicitacao", linha_guia.get("dataSolicitacao"), is_date=True)
            sub(guia, "numeroGuiaSPSADTPrincipal", linha_guia.get("numeroGuiaSPSADTPrincipal"))
            sub(guia, "dataAutorizacao", linha_guia.get("dataAutorizacao"), is_date=True)
            sub(guia, "dataRealizacao", linha_guia.get("dataRealizacao"), is_date=True)
            sub(guia, "dataInicialFaturamento", linha_guia.get("dataInicialFaturamento"), is_date=True)
            sub(guia, "dataFimPeriodo", linha_guia.get("dataFimPeriodo"), is_date=True)
            sub(guia, "dataProtocoloCobranca", linha_guia.get("dataProtocoloCobranca"), is_date=True)
            sub(guia, "dataPagamento", linha_guia.get("dataPagamento"), is_date=True)
            sub(guia, "dataProcessamentoGuia", linha_guia.get("dataProcessamentoGuia"), is_date=True)
            sub(guia, "tipoConsulta", linha_guia.get("tipoConsulta"))
            sub(guia, "cboExecutante", linha_guia.get("cboExecutante"))
            sub(guia, "indicacaoRecemNato", linha_guia.get("indicacaoRecemNato"))
            sub(guia, "indicacaoAcidente", linha_guia.get("indicacaoAcidente"))
            sub(guia, "caraterAtendimento", linha_guia.get("caraterAtendimento"))
            sub(guia, "tipoInternacao", linha_guia.get("tipoInternacao"))
            sub(guia, "regimeInternacao", linha_guia.get("regimeInternacao"))
            
            if pd.notna(linha_guia.get("diagnosticoCID")):
                diagnosticosCID10_el = ET.SubElement(guia, "ans:diagnosticosCID10")
                sub(diagnosticosCID10_el, "diagnosticoCID", linha_guia.get("diagnosticoCID"))
            
            sub(guia, "tipoAtendimento", linha_guia.get("tipoAtendimento"))
            sub(guia, "regimeAtendimento", linha_guia.get("regimeAtendimento"))
            sub(guia, "tipoFaturamento", linha_guia.get("tipoFaturamento"))
            sub(guia, "diariasAcompanhante", linha_guia.get("diariasAcompanhante"))
            sub(guia, "diariasUTI", linha_guia.get("diariasUTI"))
            sub(guia, "motivoSaida", linha_guia.get("motivoSaida"))

            valoresGuia_el = ET.SubElement(guia, "ans:valoresGuia")
            sub(valoresGuia_el, "valorTotalInformado", linha_guia.get("valorTotalInformado"))
            sub(valoresGuia_el, "valorProcessado", linha_guia.get("valorProcessado"))
            sub(valoresGuia_el, "valorTotalPagoProcedimentos", linha_guia.get("valorTotalPagoProcedimentos"))
            sub(valoresGuia_el, "valorTotalDiarias", linha_guia.get("valorTotalDiarias"))
            sub(valoresGuia_el, "valorTotalTaxas", linha_guia.get("valorTotalTaxas"))
            sub(valoresGuia_el, "valorTotalMateriais", linha_guia.get("valorTotalMateriais"))
            sub(valoresGuia_el, "valorTotalOPME", linha_guia.get("valorTotalOPME"))
            sub(valoresGuia_el, "valorTotalMedicamentos", linha_guia.get("valorTotalMedicamentos"))
            sub(valoresGuia_el, "valorGlosaGuia", linha_guia.get("valorGlosaGuia"))
            sub(valoresGuia_el, "valorPagoGuia", linha_guia.get("valorPagoGuia"))
            sub(valoresGuia_el, "valorPagoFornecedores", linha_guia.get("valorPagoFornecedores"))
            sub(valoresGuia_el, "valorTotalTabelaPropria", linha_guia.get("valorTotalTabelaPropria"))
            sub(valoresGuia_el, "valorTotalCoParticipacao", linha_guia.get("valorTotalCoParticipacao"))

            sub(guia, "declaracaoNascido", linha_guia.get("declaracaoNascido"))
            sub(guia, "declaracaoObito", linha_guia.get("declaracaoObito"))

            # --- Loop Interno para cada Procedimento da Guia ---
            for _, proc_linha in grupo_guia_key.iterrows():
                if pd.notna(proc_linha.get("codigoProcedimento")) or pd.notna(proc_linha.get("grupoProcedimento")):
                    procedimentos_el = ET.SubElement(guia, "ans:procedimentos")
                    identProcedimento_el = ET.SubElement(procedimentos_el, "ans:identProcedimento")
                    sub(identProcedimento_el, "codigoTabela", proc_linha.get("codigoTabela"))
                    Procedimento_el = ET.SubElement(identProcedimento_el, "ans:Procedimento")
                    
                    if pd.notna(proc_linha.get("grupoProcedimento")):
                        sub(Procedimento_el, "grupoProcedimento", proc_linha.get("grupoProcedimento"))
                    elif pd.notna(proc_linha.get("codigoProcedimento")):
                        sub(Procedimento_el, "codigoProcedimento", proc_linha.get("codigoProcedimento"))
                    
                    sub(procedimentos_el, "quantidadeInformada", proc_linha.get("quantidadeInformada"))
                    sub(procedimentos_el, "valorInformado", proc_linha.get("valorInformado_proc"))
                    sub(procedimentos_el, "quantidadePaga", proc_linha.get("quantidadePaga"))
                    sub(procedimentos_el, "unidadeMedida", proc_linha.get("unidadeMedida"))
                    sub(procedimentos_el, "valorPagoProc", proc_linha.get("valorPagoProc"))
                    sub(procedimentos_el, "valorPagoFornecedor", proc_linha.get("valorPagoFornecedor_proc"))
                    sub(procedimentos_el, "valorCoParticipacao", proc_linha.get("valorCoParticipacao"))

        # --- Finalização com Hash e Formatação ---
        conteudo_cabecalho = ''.join(extrair_texto(cabecalho))
        conteudo_mensagem = ''.join(extrair_texto(mensagem))
        conteudo_para_hash = conteudo_cabecalho + conteudo_mensagem
        hash_value = hashlib.md5(conteudo_para_hash.encode('iso-8859-1')).hexdigest()
        epilogo = ET.SubElement(root, "ans:epilogo")
        ET.SubElement(epilogo, "ans:hash").text = hash_value
        xml_string = ET.tostring(root, encoding="utf-8", method="xml")
        dom = minidom.parseString(xml_string)
        final_pretty = dom.toprettyxml(indent="  ", encoding="iso-8859-1")
        nome_base, _ = os.path.splitext(nome_arquivo)
        nome_limpo = re.sub(r'[^a-zA-Z0-9_\-]', '_', nome_base)
        arquivos_gerados[f"{nome_limpo}.xml"] = final_pretty
        arquivos_gerados[f"{nome_limpo}.xte"] = final_pretty

    return arquivos_gerados
######################################### STREAM LIT #########################################  


# Forçar tema escuro
st.set_page_config(page_title="Conversor Avançado de XTE", layout="wide")

# Custom CSS para destaque do menu
st.markdown("""
    <style>
        section[data-testid="stSidebar"] .css-ng1t4o {
            background-color: #1e1e1e;
            color: white;
            font-weight: bold;
            font-size: 1.1rem;
        }
        section[data-testid="stSidebar"] label {
            color: white !important;
        }
    </style>
""", unsafe_allow_html=True)

st.sidebar.title("AM Consultoria")
menu = st.sidebar.radio("Escolha uma operação:", [
    "Converter XTE para Excel e CSV",
    "Converter Excel para XTE/XML"
])

st.title("Conversor Avançado de XTE ⇄ Excel")

if menu == "Converter XTE para Excel e CSV":
    st.subheader("📄➡📊 Transformar arquivos .XTE em Excel e CSV")
    
    st.markdown("""
    Este modo permite que você envie **dois ou mais arquivos `.xte`** e receba:

    - Um **arquivo Excel (.xlsx)** consolidado.
    - Um **arquivo CSV (.csv)** com os mesmos dados.

    Ideal para visualizar, editar e analisar seus dados fora do sistema.
    """)

    uploaded_files = st.file_uploader("Selecione os arquivos .xte", accept_multiple_files=True, type=["xte"])

    if uploaded_files:
        st.info(f"Você enviou {len(uploaded_files)} arquivos. Aguarde enquanto processamos.")
        progress_bar = st.progress(0)
        status_text = st.empty()
        all_dfs = []

        total = len(uploaded_files)
        start_time = time.time()

        for i, file in enumerate(uploaded_files):
            step_start = time.time()
            with st.spinner(f"Lendo arquivo {file.name}..."):
                df, _, _ = parse_xte(file)
                df['Nome da Origem'] = file.name
                all_dfs.append(df)

            elapsed = time.time() - start_time
            avg_time = elapsed / (i + 1)
            est_remaining = avg_time * (total - (i + 1))

            percent_complete = (i + 1) / total
            progress_bar.progress(percent_complete)

            status_text.markdown(
                f"Processado {i + 1} de {total} arquivos ({percent_complete:.0%})  \
                Estimado restante: {int(est_remaining)} segundos 🕒"
            )

        final_df = pd.concat(all_dfs, ignore_index=True)
        st.success(f"✅ Processamento concluído: {len(final_df)} registros.")

        st.subheader("🔍 Pré-visualização dos dados:")
        st.dataframe(final_df.head(20))

        excel_buffer = io.BytesIO()
        final_df.to_excel(excel_buffer, index=False)

        csv_buffer = io.StringIO()
        final_df.to_csv(csv_buffer, index=False, sep=";", encoding="utf-8", float_format='%.2f')

        st.download_button("⬇ Baixar Excel Consolidado", data=excel_buffer.getvalue(), file_name="dados_consolidados.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.download_button("⬇ Baixar CSV Consolidado", data=csv_buffer.getvalue(), file_name="dados_consolidados.csv", mime="text/csv")

elif menu == "Converter Excel para XTE/XML":
    st.subheader("📊➡📄 Transformar Excel em arquivos .XTE/XML")

    st.markdown("""
    Aqui você pode carregar **um arquivo Excel atualizado** e o sistema irá:

    - Processar os dados.
    - Gerar **vários arquivos `.xte` ou `.xml`**.
    - Compactar os arquivos `.xml` automaticamente.
    - Permitir que você baixe os arquivos `.xte` somente quando desejar.

    **Antes disso**, você poderá baixar **um exemplo do primeiro arquivo gerado.**
    """)

    excel_file = st.file_uploader("Selecione o arquivo Excel (.xlsx ou .csv)", type=["xlsx", "csv"])

    if excel_file:
        st.info("🔄 Processando o arquivo...")

        try:
            with st.spinner("Gerando arquivos..."):
                updated_files = gerar_xte_do_excel(excel_file)

            # Separar XMLs e XTEs
            xml_files = {k: v for k, v in updated_files.items() if k.endswith(".xml")}
            xte_files = {k.replace(".xml", ".xte"): v for k, v in updated_files.items() if k.endswith(".xml")}

            # Exemplo de preview
            first_key = next(iter(xml_files))
            first_file = xml_files[first_key]

            st.download_button(
                f"⬇ Baixar exemplo: {first_key}",
                data=first_file,
                file_name=first_key,
                mime="application/xml"
            )

            st.download_button(
                f"⬇ Baixar exemplo em XTE: {first_key.replace('.xml', '.xte')}",
                data=first_file,
                file_name=first_key.replace('.xml', '.xte'),
                mime="application/xml"
            )

            # Compactar XMLs automaticamente
            st.info("📦 Compactando arquivos XML...")
            xml_zip_buffer = io.BytesIO()
            start_time = time.time()
            progress = st.progress(0)
            status = st.empty()
            total = len(xml_files)

            with zipfile.ZipFile(xml_zip_buffer, "w") as zipf:
                for i, (filename, content) in enumerate(xml_files.items()):
                    zipf.writestr(filename, content)
                    elapsed = time.time() - start_time
                    avg = elapsed / (i + 1)
                    remaining = avg * (total - (i + 1))
                    progress.progress((i + 1) / total)
                    status.markdown(f"📄 Adicionando {i + 1}/{total} arquivos XML - ⏳ Restante: {int(remaining)}s")

            st.success("✅ Arquivo ZIP com XMLs pronto!")
            st.download_button(
                "⬇ Baixar ZIP de XMLs",
                data=xml_zip_buffer.getvalue(),
                file_name="arquivos_xml.zip",
                mime="application/zip"
            )

            # Botão para gerar e baixar XTEs
            if st.button("📁 Gerar e Baixar Arquivo ZIP com XTEs"):
                st.info("📦 Compactando arquivos XTE...")
                xte_zip_buffer = io.BytesIO()
                start_time = time.time()
                progress_xte = st.progress(0)
                status_xte = st.empty()
                total_xte = len(xte_files)

                with zipfile.ZipFile(xte_zip_buffer, "w") as zipf:
                    for i, (filename, content) in enumerate(xte_files.items()):
                        zipf.writestr(filename, content)
                        elapsed = time.time() - start_time
                        avg = elapsed / (i + 1)
                        remaining = avg * (total_xte - (i + 1))
                        progress_xte.progress((i + 1) / total_xte)
                        status_xte.markdown(f"📄 Adicionando {i + 1}/{total_xte} arquivos XTE - ⏳ Restante: {int(remaining)}s")

                st.success("✅ Arquivo ZIP com XTEs pronto!")
                st.download_button(
                    "⬇ Baixar ZIP de XTEs",
                    data=xte_zip_buffer.getvalue(),
                    file_name="arquivos_xte.zip",
                    mime="application/zip"
                )

        except Exception as e:
            st.error(f"Erro durante o processamento: {str(e)}")
            st.error("Verifique se o arquivo Excel possui a estrutura correta.")
