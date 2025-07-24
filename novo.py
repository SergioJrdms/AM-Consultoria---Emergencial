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
import pytz

# --- CONSTANTES GLOBAIS PARA A NOVA VERSÃO ---
TISS_VERSAO = "5.01.00"
TISS_NAMESPACE = "http://www.ans.gov.br/padroes/tiss/schemas"
TISS_SCHEMA_FILE = f"tissMonitoramentoV{TISS_VERSAO.replace('.', '_')}.xsd"
TISS_SCHEMA_LOCATION = f"{TISS_NAMESPACE} {TISS_NAMESPACE}/{TISS_SCHEMA_FILE}"

@st.cache_data
def parse_xte(file):
    """
    Função atualizada para ler arquivos .xte no novo formato TISS 5.01.00.
    """
    file.seek(0)
    content = file.read().decode('iso-8859-1')
    tree = ET.ElementTree(ET.fromstring(content))
    root = tree.getroot()
    ns = {'ans': TISS_NAMESPACE}
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
        cabecalho_info['registroANS_cabecalho'] = cabecalho.findtext('ans:registroANS', default='', namespaces=ns)
        cabecalho_info['versaoPadrao_cabecalho'] = cabecalho.findtext('ans:versaoPadrao', default='', namespaces=ns)

    # A tag principal agora é monitoramentoSaudeSuplementar
    for guia in root.findall(".//ans:monitoramentoSaudeSuplementar", namespaces=ns):
        guia_data = cabecalho_info.copy()

        # Extração de dados da nova estrutura
        # Mapeia diretamente os campos para evitar complexidade
        campos_guia = {
            'identificacaoMonitorado/registroANS': 'registroANS_monitorado',
            'identificacaoMonitorado/cnpjOperadora': 'cnpjOperadora',
            'identificacaoMonitorado/dataEmissao': 'dataEmissao',
            'dadosBeneficiario/numeroCarteira': 'numeroCarteira',
            'dadosBeneficiario/tempoPlano': 'tempoPlano',
            'dadosBeneficiario/nomeBeneficiario': 'nomeBeneficiario',
            'dadosBeneficiario/dataNascimento': 'dataNascimento',
            'dadosBeneficiario/sexo': 'sexo',
            'dadosBeneficiario/codigoMunicipio': 'codigoMunicipioBeneficiario',
            'dadosBeneficiario/numeroContrato': 'numeroContrato',
            'dadosBeneficiario/tipoPlano': 'tipoPlano',
            'dadosContratado/identificacao/codigoNaOperadora': 'codigoContratadoNaOperadora',
            'dadosContratado/identificacao/cpf': 'cpfContratado',
            'dadosContratado/identificacao/cnpj': 'cnpjContratado',
            'dadosContratado/nomeContratado': 'nomeContratado',
            'eventosAtencaoSaude/numeroGuiaPrestador': 'numeroGuiaPrestador',
            'eventosAtencaoSaude/numeroGuiaOperadora': 'numeroGuiaOperadora',
            'eventosAtencaoSaude/senha': 'senha',
            'eventosAtencaoSaude/tipoAtendimento': 'tipoAtendimento',
            'eventosAtencaoSaude/indicadorRecemNascido': 'indicadorRecemNascido',
            'eventosAtencaoSaude/indicadorAcidente': 'indicadorAcidente',
            'eventosAtencaoSaude/dataRealizacao': 'dataRealizacao',
            'eventosAtencaoSaude/caraterAtendimento': 'caraterAtendimento',
            'eventosAtencaoSaude/cboProfissional': 'cboProfissional',
            'totaisGuia/valorTotalInformado': 'valorTotalInformado',
            'totaisGuia/valorTotalProcessado': 'valorTotalProcessado',
            'totaisGuia/valorTotalLiberado': 'valorTotalLiberado',
            'totaisGuia/valorTotalGlosa': 'valorTotalGlosa',
        }

        for path, col_name in campos_guia.items():
            valor = guia.findtext(f'ans:{path.replace("/", "/ans:")}', default=None, namespaces=ns)
            guia_data[col_name] = valor.strip() if valor else None

        procedimentos = guia.findall(".//ans:procedimentosRealizados", namespaces=ns)
        if procedimentos:
            for proc in procedimentos:
                proc_data = guia_data.copy()
                # Extração específica dos procedimentos
                proc_data['codigoTabela'] = proc.findtext('ans:codigoTabela', default='', namespaces=ns).strip()
                proc_data['codigoProcedimento'] = proc.findtext('ans:codigoProcedimento', default='', namespaces=ns).strip()
                proc_data['descricaoProcedimento'] = proc.findtext('ans:descricaoProcedimento', default='', namespaces=ns).strip()
                proc_data['quantidadeExecutada'] = proc.findtext('ans:quantidadeExecutada', default='', namespaces=ns).strip()
                proc_data['valorInformado'] = proc.findtext('ans:valorInformado', default='', namespaces=ns).strip()
                proc_data['valorProcessado'] = proc.findtext('ans:valorProcessado', default='', namespaces=ns).strip()
                proc_data['valorLiberado'] = proc.findtext('ans:valorLiberado', default='', namespaces=ns).strip()
                proc_data['valorGlosa'] = proc.findtext('ans:valorGlosa', default='', namespaces=ns).strip()

                all_data.append(proc_data)
        else:
            all_data.append(guia_data)

    df = pd.DataFrame(all_data)
    df['Nome da Origem'] = file.name

    date_columns = [col for col in df.columns if 'data' in col.lower()]
    for col in date_columns:
        try:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d/%m/%Y')
        except Exception:
            pass

    # Calcular idade
    if 'dataRealizacao' in df.columns and 'dataNascimento' in df.columns:
        def calcular_idade(row):
            try:
                data_realizacao = datetime.strptime(row['dataRealizacao'], '%d/%m/%Y')
                data_nascimento = datetime.strptime(row['dataNascimento'], '%d/%m/%Y')
                return (data_realizacao - data_nascimento).days // 365
            except (ValueError, TypeError):
                return None
        df['Idade_na_Realizacao'] = df.apply(calcular_idade, axis=1)

    # Lista de colunas final baseada na nova estrutura TISS 5.01.00
    colunas_finais = [
        'Nome da Origem', 'tipoTransacao', 'numeroLote', 'competenciaLote', 'dataRegistroTransacao', 'horaRegistroTransacao',
        'registroANS_cabecalho', 'versaoPadrao_cabecalho', 'registroANS_monitorado', 'cnpjOperadora', 'dataEmissao',
        'numeroCarteira', 'tempoPlano', 'nomeBeneficiario', 'dataNascimento', 'sexo', 'codigoMunicipioBeneficiario',
        'numeroContrato', 'tipoPlano', 'codigoContratadoNaOperadora', 'cpfContratado', 'cnpjContratado', 'nomeContratado',
        'numeroGuia_prestador', 'numeroGuiaOperadora', 'senha', 'tipoAtendimento', 'indicadorRecemNascido',
        'indicadorAcidente', 'dataRealizacao', 'caraterAtendimento', 'cboProfissional',
        'valorTotalInformado', 'valorTotalProcessado', 'valorTotalLiberado', 'valorTotalGlosa',
        'codigoTabela', 'codigoProcedimento', 'descricaoProcedimento', 'quantidadeExecutada', 'valorInformado',
        'valorProcessado', 'valorLiberado', 'valorGlosa', 'Idade_na_Realizacao'
    ]

    for col in colunas_finais:
        if col not in df.columns:
            df[col] = None

    return df[colunas_finais], content, tree

def gerar_xte_do_excel(excel_file):
    """
    Função atualizada para gerar arquivos .xte no novo formato TISS 5.01.00.
    """
    print(f"--- DEBUG: Gerando XTE no padrão TISS {TISS_VERSAO} ---")

    # --- Setup de Data/Hora e Leitura do Arquivo ---
    fuso_horario_servidor = pytz.utc
    fuso_horario_desejado = pytz.timezone("America/Sao_Paulo")
    agora_no_fuso_desejado = datetime.now(fuso_horario_servidor).astimezone(fuso_horario_desejado)
    data_atual = agora_no_fuso_desejado.strftime("%Y-%m-%d")
    hora_atual = agora_no_fuso_desejado.strftime("%H:%M:%S")
    minuto_e_segundos_atuais = agora_no_fuso_desejado.strftime("%M%S")

    if hasattr(excel_file, 'name') and excel_file.name.endswith('.csv'):
        df = pd.read_csv(excel_file, dtype=str, sep=';')
    else:
        df = pd.read_excel(excel_file, dtype=str)

    # --- Função Auxiliar 'sub' ---
    def sub(parent, tag, value, is_date=False):
        if pd.isna(value) or str(value).strip() == '':
            return
        text = str(value).strip()
        if is_date and text:
            try:
                # Tenta converter de D/M/A para A-M-D
                text = datetime.strptime(text.split(' ')[0], "%d/%m/%Y").strftime("%Y-%m-%d")
            except ValueError:
                # Assume que já está no formato correto se a conversão falhar
                pass
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

        agrupado = df_origem.groupby(["numeroGuia_prestador"], dropna=False)

        root = ET.Element("ans:mensagemTISS", attrib={
            "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance", "xmlns:xsd": "http://www.w3.org/2001/XMLSchema",
            "xsi:schemaLocation": TISS_SCHEMA_LOCATION, "xmlns:ans": TISS_NAMESPACE
        })

        linha_cabecalho = df_origem.iloc[0]

        # --- Bloco do Cabeçalho ---
        cabecalho = ET.SubElement(root, "ans:cabecalho")
        identificacaoTransacao = ET.SubElement(cabecalho, "ans:identificacaoTransacao")
        sub(identificacaoTransacao, "tipoTransacao", "MONITORAMENTO_SAUDE_SUPLEMENTAR")

        competencia = linha_cabecalho.get("competenciaLote", "")
        if competencia and len(competencia) == 6 and competencia.isdigit():
            numero_lote_final = f"{competencia}{minuto_e_segundos_atuais}"
        else:
            ano_e_mes_atuais = agora_no_fuso_desejado.strftime("%Y%m")
            numero_lote_final = f"{ano_e_mes_atuais}{minuto_e_segundos_atuais}"

        sub(identificacaoTransacao, "numeroLote", numero_lote_final)
        sub(identificacaoTransacao, "competenciaLote", competencia)
        sub(identificacaoTransacao, "dataRegistroTransacao", data_atual)
        sub(identificacaoTransacao, "horaRegistroTransacao", hora_atual)
        sub(cabecalho, "registroANS", linha_cabecalho.get("registroANS_cabecalho"))
        sub(cabecalho, "versaoPadrao", TISS_VERSAO)

        mensagem = ET.SubElement(root, "ans:mensagem")
        op_ans = ET.SubElement(mensagem, "ans:operadoraParaANS")

        # --- Loop Principal para cada Guia (agora registro de monitoramento) ---
        for _, grupo_guia_key in agrupado:
            linha_guia = grupo_guia_key.iloc[0]
            monitoramento = ET.SubElement(op_ans, "ans:monitoramentoSaudeSuplementar")

            # --- Mapeamento da nova estrutura 5.01.00 ---
            ident_monitorado = ET.SubElement(monitoramento, "ans:identificacaoMonitorado")
            sub(ident_monitorado, "registroANS", linha_guia.get("registroANS_monitorado"))
            sub(ident_monitorado, "cnpjOperadora", linha_guia.get("cnpjOperadora"))
            sub(ident_monitorado, "dataEmissao", data_atual, is_date=True)

            dados_benef = ET.SubElement(monitoramento, "ans:dadosBeneficiario")
            sub(dados_benef, "numeroCarteira", linha_guia.get("numeroCarteira"))
            sub(dados_benef, "tempoPlano", linha_guia.get("tempoPlano"))
            sub(dados_benef, "nomeBeneficiario", linha_guia.get("nomeBeneficiario"))
            sub(dados_benef, "dataNascimento", linha_guia.get("dataNascimento"), is_date=True)
            sub(dados_benef, "sexo", linha_guia.get("sexo"))
            sub(dados_benef, "codigoMunicipio", linha_guia.get("codigoMunicipioBeneficiario"))
            sub(dados_benef, "numeroContrato", linha_guia.get("numeroContrato"))
            sub(dados_benef, "tipoPlano", linha_guia.get("tipoPlano"))

            dados_contratado = ET.SubElement(monitoramento, "ans:dadosContratado")
            ident_contratado = ET.SubElement(dados_contratado, "ans:identificacao")
            sub(ident_contratado, "codigoNaOperadora", linha_guia.get("codigoContratadoNaOperadora"))
            # Adiciona CPF ou CNPJ conforme o que estiver preenchido
            if pd.notna(linha_guia.get("cpfContratado")):
                 ET.SubElement(ident_contratado, "ans:cpf").text = linha_guia.get("cpfContratado")
            elif pd.notna(linha_guia.get("cnpjContratado")):
                 ET.SubElement(ident_contratado, "ans:cnpj").text = linha_guia.get("cnpjContratado")
            sub(dados_contratado, "nomeContratado", linha_guia.get("nomeContratado"))

            eventos = ET.SubElement(monitoramento, "ans:eventosAtencaoSaude")
            sub(eventos, "numeroGuiaPrestador", linha_guia.get("numeroGuia_prestador"))
            sub(eventos, "numeroGuiaOperadora", linha_guia.get("numeroGuiaOperadora"))
            sub(eventos, "senha", linha_guia.get("senha"))
            sub(eventos, "tipoAtendimento", linha_guia.get("tipoAtendimento"))
            sub(eventos, "indicadorRecemNascido", linha_guia.get("indicadorRecemNascido"))
            sub(eventos, "indicadorAcidente", linha_guia.get("indicadorAcidente"))
            sub(eventos, "dataRealizacao", linha_guia.get("dataRealizacao"), is_date=True)
            sub(eventos, "caraterAtendimento", linha_guia.get("caraterAtendimento"))
            sub(eventos, "cboProfissional", linha_guia.get("cboProfissional"))

            # --- Loop Interno para cada Procedimento da Guia ---
            for _, proc_linha in grupo_guia_key.iterrows():
                if pd.notna(proc_linha.get("codigoProcedimento")):
                    proc_realizado = ET.SubElement(eventos, "ans:procedimentosRealizados")
                    sub(proc_realizado, "codigoTabela", proc_linha.get("codigoTabela"))
                    sub(proc_realizado, "codigoProcedimento", proc_linha.get("codigoProcedimento"))
                    sub(proc_realizado, "descricaoProcedimento", proc_linha.get("descricaoProcedimento"))
                    sub(proc_realizado, "quantidadeExecutada", proc_linha.get("quantidadeExecutada"))
                    sub(proc_realizado, "valorInformado", proc_linha.get("valorInformado"))
                    sub(proc_realizado, "valorProcessado", proc_linha.get("valorProcessado"))
                    sub(proc_realizado, "valorLiberado", proc_linha.get("valorLiberado"))
                    sub(proc_realizado, "valorGlosa", proc_linha.get("valorGlosa"))

            totais_guia = ET.SubElement(monitoramento, "ans:totaisGuia")
            sub(totais_guia, "valorTotalInformado", linha_guia.get("valorTotalInformado"))
            sub(totais_guia, "valorTotalProcessado", linha_guia.get("valorTotalProcessado"))
            sub(totais_guia, "valorTotalLiberado", linha_guia.get("valorTotalLiberado"))
            sub(totais_guia, "valorTotalGlosa", linha_guia.get("valorTotalGlosa"))

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
        arquivos_gerados[f"{nome_limpo}.xte"] = final_pretty # XTE tem o mesmo conteúdo do XML

    return arquivos_gerados

######################################### STREAMLIT UI (sem grandes alterações) #########################################

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
st.sidebar.markdown(f"**Padrão TISS:** `{TISS_VERSAO}`")
menu = st.sidebar.radio("Escolha uma operação:", [
    "Converter XTE para Excel e CSV",
    "Converter Excel para XTE/XML"
])

st.title("Conversor Avançado de XTE ⇄ Excel")
st.warning(f":information_source: **Atenção:** Este conversor foi atualizado para o Padrão TISS versão **{TISS_VERSAO}**, com vigência a partir de **01/03/2025**.")


if menu == "Converter XTE para Excel e CSV":
    st.subheader(":page_facing_up:➡:bar_chart: Transformar arquivos .XTE em Excel e CSV")

    st.markdown("""
    Este modo permite que você envie **um ou mais arquivos `.xte`** (no formato TISS 5.01.00) e receba:

    - Um **arquivo Excel (.xlsx)** consolidado.
    - Um **arquivo CSV (.csv)** com os mesmos dados.

    Ideal para visualizar, editar e analisar seus dados fora do sistema.
    """)

    uploaded_files = st.file_uploader("Selecione os arquivos .xte", accept_multiple_files=True, type=["xte", "xml"])

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
                try:
                    df, _, _ = parse_xte(file)
                    df['Nome da Origem'] = file.name
                    all_dfs.append(df)
                except Exception as e:
                    st.error(f"Erro ao processar o arquivo {file.name}: {e}")
                    st.error("Verifique se o arquivo está no formato TISS 5.01.00 correto.")
                    continue

            elapsed = time.time() - start_time
            avg_time = elapsed / (i + 1)
            est_remaining = avg_time * (total - (i + 1))

            percent_complete = (i + 1) / total
            progress_bar.progress(percent_complete)

            status_text.markdown(
                f"Processado {i + 1} de {total} arquivos ({percent_complete:.0%})  \
                Estimado restante: {int(est_remaining)} segundos :clock3:"
            )
        
        if all_dfs:
            final_df = pd.concat(all_dfs, ignore_index=True)
            st.success(f":white_check_mark: Processamento concluído: {len(final_df)} registros.")

            st.subheader(":mag: Pré-visualização dos dados (formato TISS 5.01.00):")
            st.dataframe(final_df.head(20))

            excel_buffer = io.BytesIO()
            final_df.to_excel(excel_buffer, index=False)

            csv_buffer = io.StringIO()
            final_df.to_csv(csv_buffer, index=False, sep=";", encoding="utf-8-sig")

            st.download_button("⬇ Baixar Excel Consolidado", data=excel_buffer.getvalue(), file_name="dados_consolidados_tiss5.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.download_button("⬇ Baixar CSV Consolidado", data=csv_buffer.getvalue(), file_name="dados_consolidados_tiss5.csv", mime="text/csv")
        else:
            st.warning("Nenhum dado foi processado. Verifique os arquivos.")


elif menu == "Converter Excel para XTE/XML":
    st.subheader(":bar_chart:➡:page_facing_up: Transformar Excel em arquivos .XTE/XML")

    st.markdown("""
    Aqui você pode carregar **um arquivo Excel ou CSV** com os dados no leiaute TISS 5.01.00 e o sistema irá:

    - Processar os dados.
    - Gerar **vários arquivos `.xte` e `.xml`**.
    - Compactar os arquivos `.xml` e `.xte` em arquivos ZIP separados.

    **Recomendação:** Use a função de 'Converter XTE para Excel' para gerar um modelo em branco com as colunas corretas para preenchimento.
    """)

    excel_file = st.file_uploader("Selecione o arquivo Excel (.xlsx ou .csv)", type=["xlsx", "csv"])

    if excel_file:
        st.info(":arrows_counterclockwise: Processando o arquivo...")

        try:
            with st.spinner("Gerando arquivos no formato TISS 5.01.00..."):
                updated_files = gerar_xte_do_excel(excel_file)

            if not updated_files:
                st.error("Nenhum arquivo foi gerado. Verifique se o Excel contém dados e a coluna 'Nome da Origem'.")
            else:
                st.success(f":tada: Sucesso! {len(updated_files) // 2} arquivos XML/XTE foram gerados.")

                xml_files = {k: v for k, v in updated_files.items() if k.endswith(".xml")}
                xte_files = {k: v for k, v in updated_files.items() if k.endswith(".xte")}

                # Preview do primeiro arquivo
                first_key = next(iter(xml_files))
                first_file_content = xml_files[first_key]

                st.download_button(
                    f"⬇ Baixar exemplo XML: {first_key}",
                    data=first_file_content,
                    file_name=first_key,
                    mime="application/xml"
                )
                
                # Botão para baixar todos os XMLs
                xml_zip_buffer = io.BytesIO()
                with zipfile.ZipFile(xml_zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                    for filename, content in xml_files.items():
                        zipf.writestr(filename, content)
                st.download_button(
                    "⬇ Baixar TODOS os arquivos .XML (.zip)",
                    data=xml_zip_buffer.getvalue(),
                    file_name="arquivos_xml_tiss5.zip",
                    mime="application/zip"
                )

                # Botão para baixar todos os XTEs
                xte_zip_buffer = io.BytesIO()
                with zipfile.ZipFile(xte_zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                    for filename, content in xte_files.items():
                        zipf.writestr(filename, content)
                st.download_button(
                    "⬇ Baixar TODOS os arquivos .XTE (.zip)",
                    data=xte_zip_buffer.getvalue(),
                    file_name="arquivos_xte_tiss5.zip",
                    mime="application/zip"
                )

        except Exception as e:
            st.error(f"Ocorreu um erro durante o processamento: {str(e)}")
            st.error("Por favor, verifique se a estrutura do seu arquivo Excel/CSV corresponde ao novo padrão TISS 5.01.00.")
