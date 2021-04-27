import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import base64
from io import BytesIO

st.set_option('deprecation.showPyplotGlobalUse', False)


# Exportando os relatórios Excel
def to_excel(df, df1, df2, df3, df4, df5, df6):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Links', index=False)
    df1.to_excel(writer, sheet_name='Escopo_Links', index=False)
    df2.to_excel(writer, sheet_name='Equipamentos por Links', index=False)
    df3.to_excel(writer, sheet_name='Resumo_IDUs', index=False)
    df4.to_excel(writer, sheet_name='Resumo_ODUs', index=False)
    df5.to_excel(writer, sheet_name='Resumo_Antenas', index=False)
    df6.to_excel(writer, sheet_name='Resumo_Config', index=False)
    writer.save()
    processed_data = output.getvalue()
    return processed_data


def get_table_download_link(df, df1, df2, df3, df4, df5, df6):
    """Generates a link allowing the data in a given panda dataframe to be downloaded
    in:  dataframe
    out: href string
    """
    val = to_excel(df, df1, df2, df3, df4, df5, df6)
    b64 = base64.b64encode(val)  # val looks like b'...'
    return f'<a href="data:application/octet-stream;base64,{b64.decode()}' \
           f'" download="Export_forecast.xlsx">Download file</a>'


@st.cache
def load_data(file):
    """Function for loading data"""
    df_file = pd.read_excel(file)
    return df_file


# title for the app
st.title("Forecast App")

# Setup file upload
st.sidebar.subheader('Import from Excel')
uploaded_file = st.sidebar.file_uploader(label='Upload your Excel file',
                                         type=['csv', 'xlsm', 'xlsx'])
global df
if uploaded_file is not None:
    df = load_data(uploaded_file)

    # Ajustando o Data Frame original
    df = df.rename(columns={'REUSO _ CH1 - SBB': 'REUSO_CH1 - SBB', })
    df['Escopo'] = df['Escopo'].str.upper()
    df['Status Instalacao'] = df['Status Instalacao'].str.upper()
    df['Planning Status'] = df['Planning Status'].str.upper()
    df['Status Project'] = df['Status Project'].str.upper()
    df['Link Type Engenharia'] = df['Link Type Engenharia'].str.upper()
    df = df.astype({'ANT A': 'float', 'ANT B': 'float', 'ANT SD A': 'float', 'ANT SD B': 'float'})
    for i in df[df['HW A - MODEL'] == 'AGS20L']['IDU A - GAI'].index:
        df.loc[i, 'IDU A - GAI'] = 'AGS20L'
        df.loc[i, 'IDU B - GAI'] = 'AGS20L'

    # Filtro retirando links Instalados, Cancelados e Reuso
    df1 = df[(df['Status Instalacao'] == "PENDENTE") & (df['Planning Status'] == "PLANNED")
             & (df['Status Project'] == "ON PROJECT") & (df['Link Type Engenharia'] != "REUSO")]
    df1 = df1.fillna('')

    # Analise escopo Links
    count_escopo = df1.groupby(['Regional', 'Escopo',
                                'Status Instalacao']).size().reset_index(name='Count')

    # Analise config Links
    count_config = df1.groupby('CONF').size().reset_index(name='Count')

    # Sinalizando Links não analisados
    df['Analisado'] = df['SIAE ID'].isin(df1['SIAE ID'])
    links_analisados = df['Analisado'].value_counts()

    # Preparando o DataFrames para análise de odus
    ajuste_odu = {'1.0': '1', '2.0': '2', '3.0': '3', '4.0': '4', '1 or 2': '1', '2 or 3': '2',
                  '1 or 4': '1', '2 or 4': '2', '3 or 4': '3', 'ch 1': 'ch1', 'ch 2': 'ch2', 'ch 3': 'ch3',
                  'ch 4': 'ch4', 'ch 5': 'ch5', 'ch 6': 'ch6', 'ch 7': 'ch7', 'ch 8': 'ch8'}

    for i in range(1, 9):
        df1[f'CH{i} - SBB'] = df1[f'CH{i} - SBB'].astype('str')
        df1[f'CH{i} - SBB'] = df1[f'CH{i} - SBB'].replace(ajuste_odu)

    df1['ODU TYPE'] = df1['ODU TYPE'].replace({'ASN/ASNK': 'ASNK'})

    # Analisando cada linha do data frame
    for i in df1.index:

        # Info ODUs
        odus = []
        qtd_canais = 0

        for n in range(8):

            if df1.loc[i, f'REUSO_CH{n + 1} - SBB'] == 'NO':
                odus.append(df1.loc[i, f'CH{n + 1} - SBB'])
            else:
                odus.append('')

            if df1.loc[i, f'CH{n + 1} - TXA'] != '':
                qtd_canais += 1

        lista_check = ['5', '6L', '6U', '7', '8', '8.5', '7/8', '7/8.5', '8/8.5']

        if df1.loc[i, 'BW'] == 56 and df1.loc[i, 'FREQUENCY'] in lista_check:
            if qtd_canais > int(df1.loc[i, 'CONF'][0]):
                odus[1] = ''
                odus[3] = ''
                odus[5] = ''
                odus[7] = ''

        df1.loc[i, 'Sbb1'] = odus.count('1')
        df1.loc[i, 'Sbb2'] = odus.count('2')
        df1.loc[i, 'Sbb3'] = odus.count('3')
        df1.loc[i, 'Sbb4'] = odus.count('4')
        df1.loc[i, 'ch1'] = odus.count('ch1')
        df1.loc[i, 'ch2'] = odus.count('ch2')
        df1.loc[i, 'ch3'] = odus.count('ch3')
        df1.loc[i, 'ch4'] = odus.count('ch4')
        df1.loc[i, 'ch5'] = odus.count('ch5')
        df1.loc[i, 'ch6'] = odus.count('ch6')
        df1.loc[i, 'ch7'] = odus.count('ch7')
        df1.loc[i, 'ch8'] = odus.count('ch8')

        # Info Antenas
        antenas = []
        lista_str = ['A', 'B', 'SD A', 'SD B']
        for n in range(4):

            # Substitui diâmetro 0.9 para 1m:
            if df1.loc[i, f'ANT {lista_str[n]}'] == 0.9:
                df1.loc[i, f'ANT {lista_str[n]}'] = 1

            if df1.loc[i, f'REUSO _ ANT {lista_str[n][-1]}'] == 'NO':
                antenas.append(df1.loc[i, f'ANT {lista_str[n]}'])
            else:
                antenas.append('')

        df1.loc[i, '0.3'] = antenas.count(0.3)
        df1.loc[i, '0.6'] = antenas.count(0.6)
        df1.loc[i, '1'] = antenas.count(1)
        df1.loc[i, '1.2'] = antenas.count(1.2)
        df1.loc[i, '1.8'] = antenas.count(1.8)
        df1.loc[i, '2.4'] = antenas.count(2.4)
        df1.loc[i, '3'] = antenas.count(3)
        df1.loc[i, '3.6'] = antenas.count(3.6)

        # Info IDUs
        idus = []
        lista_idu_str = ['A', 'B']

        for n in range(2):

            if df1.loc[i, f'REUSO _ IDU {lista_idu_str[n]} - GAI'] == 'NO':
                idus.append(df1.loc[i, f'IDU {lista_idu_str[n]} - GAI'])
            else:
                idus.append('')

        for idu in sorted(df1['IDU A - GAI'].unique()):
            df1.loc[i, idu] = idus.count(idu)

    # Resumo qtde de Idus
    idu_A = df1[df1['REUSO _ IDU A - GAI'] == "NO"][{'HW A - MODEL', 'IDU A - GAI'}]
    idu_B = df1[df1['REUSO _ IDU B - GAI'] == "NO"][{'HW B - MODEL', 'IDU B - GAI'}]
    idu_A = idu_A.rename(columns={'HW A - MODEL': 'MODEL', 'IDU A - GAI': 'GAI'})
    idu_B = idu_B.rename(columns={'HW B - MODEL': 'MODEL', 'IDU B - GAI': 'GAI'})
    idus = pd.concat([idu_A, idu_B])
    resumo_idus = idus.groupby(['MODEL', 'GAI']).size().reset_index(name='Count')

    # Resumo qtde de ODUs
    df_odus = df1.copy()
    resumo_odus = df_odus.groupby(['FREQUENCY', 'ODU TYPE'])[
        ['Sbb1', 'Sbb2', 'Sbb3', 'Sbb4', 'ch1', 'ch2', 'ch3', 'ch4',
         'ch5', 'ch6', 'ch7', 'ch8']].sum().reset_index()
    resumo_odus = resumo_odus.drop(resumo_odus[resumo_odus.iloc[:, 2:].sum(axis=1) == 0].index)
    resumo_odus.iloc[:, 2:] = resumo_odus.iloc[:, 2:].astype('int')
    resumo_odus['Total'] = resumo_odus.sum(axis=1)

    # Resumo qtde de Antenas
    df_antenas = df1.copy()
    df_antenas['FREQUENCY'] = df_antenas['FREQUENCY'].replace(
        {"8": "7", "8.5": "7", "7/8": "7", "7/8.5": "7", "8/8.5": "7"})
    resumo_antenas = df_antenas.groupby(['FREQUENCY', 'ANT TYPE'])[
        ['0.3', '0.6', '1', '1.2', '1.8', '2.4', '3', '3.6']].sum().reset_index()
    resumo_antenas = resumo_antenas.drop(resumo_antenas[resumo_antenas.iloc[:, 2:].sum(axis=1) == 0].index)
    resumo_antenas.iloc[:, 2:] = resumo_antenas.iloc[:, 2:].astype('int')
    resumo_antenas['Total'] = resumo_antenas.sum(axis=1)

    # Resumo de materiais por ID
    equipamento_link = df1.copy()
    lista_idus = sorted(df1['IDU A - GAI'].unique())
    filtro = ['WBS', 'SIAE ID', 'Escopo', 'CONF', 'FREQUENCY', 'ODU TYPE', 'Sbb1', 'Sbb2', 'Sbb3', 'Sbb4', 'ch1', 'ch2',
              'ch3', 'ch4',
              'ch5', 'ch6', 'ch7', 'ch8', 'ANT TYPE', '0.3', '0.6', '1', '1.2', '1.8', '2.4', '3', '3.6']
    filtro.extend(lista_idus)
    equipamento_link = equipamento_link[filtro]

    # Conf Sidebar
    st.sidebar.subheader("Export to Excel")

    # Gerando o link para download do relatório
    st.sidebar.markdown(get_table_download_link(df, count_escopo, equipamento_link, resumo_idus, resumo_odus,
                                                resumo_antenas, count_config), unsafe_allow_html=True)

    # Links analisados
    st.subheader(f'Total de {len(df)} links carregados.')
    st.write(f'* {links_analisados[True]} links analisados.')
    if len(links_analisados) > 1:
        st.write(f'* {links_analisados[False]} links não analisados.')
        st.write('Links com Status de instalação finalizado, links com status de planejamento'
                 ' cancelado e links com escopo de Reuso são desconsiderados da análise final.')
    if st.checkbox('Show Raw Data'):
        st.write(df)

    # Links
    st.subheader('Resumo Links:')
    st.write('Escopo dos links divididos por Regional.')
    # Plot Escopo Links
    sns.set_theme(style="darkgrid")
    sns.catplot(x="Regional", y="Count", hue="Escopo", data=count_escopo,
                kind="bar", palette='deep', legend_out=False)
    plt.legend(loc='upper right', fontsize='xx-small')
    st.pyplot()

    if st.checkbox('Show Link Table'):
        st.write(count_escopo)

    # Plot Config Links
    sns.set_theme(style="darkgrid")
    sns.catplot(x="CONF", y="Count", data=count_config,
                kind="bar", palette='deep', legend_out=False)
    st.pyplot()

    if st.checkbox('Show Config Table'):
        st.write(count_config)

    # Antenas
    st.subheader('Resumo Antenas:')
    st.write('Quantidade total de antenas por frequência e diâmetro.')
    plot_antenas = resumo_antenas.drop('Total', axis=1)
    new_ant = plot_antenas.melt(['FREQUENCY', 'ANT TYPE'], var_name='Diameter', value_name='Count')
    sns.catplot(x="FREQUENCY", y="Count", hue="Diameter", data=new_ant,
                kind="bar", palette='deep', legend_out=False)
    plt.legend(loc='upper right', fontsize='x-small')
    st.pyplot()

    if st.checkbox('Show Antenna Table'):
        st.write(resumo_antenas)

    # ODUs
    st.subheader('Resumo ODUs:')
    st.write('Quantidade total de ODUs por frequência e SBB para cada modelo.')
    plot_odus = resumo_odus.drop('Total', axis=1)
    new_odus = plot_odus.melt(['FREQUENCY', 'ODU TYPE'], var_name='SubBanda', value_name='Count')

    # Plot ASNK
    asnk = new_odus[(new_odus['ODU TYPE'] == "ASNK") & (new_odus['Count'] != 0)]
    if len(asnk) >= 1:
        sns.catplot(x="FREQUENCY", y="Count", hue="SubBanda", data=asnk, kind="bar",
                    legend_out=False, palette='deep').set(title='ODU TYPE = ASNK')
        plt.legend(loc='upper right', fontsize='x-small')
        st.pyplot()

    # Plot ASNKHP
    asnkhp = new_odus[(new_odus['ODU TYPE'] == "ASNKHP") & (new_odus['Count'] != 0)]
    if len(asnkhp) >= 1:
        sns.catplot(x="FREQUENCY", y="Count", hue="SubBanda", data=asnkhp, kind="bar",
                    legend_out=False, palette='deep').set(title='ODU TYPE = ASNKHP')
        plt.legend(loc='upper right', fontsize='x-small')
        st.pyplot()

    # Plot ASN
    asn = new_odus[(new_odus['ODU TYPE'] == "ASN") & (new_odus['Count'] != 0)]
    if len(asn) >= 1:
        sns.catplot(x="FREQUENCY", y="Count", hue="SubBanda", data=asn, kind="bar",
                    legend_out=False, palette='deep').set(title='ODU TYPE = ASN')
        plt.legend(loc='upper right', fontsize='x-small')
        st.pyplot()

    if st.checkbox('Show ODUs Table'):
        st.write(resumo_odus)

    # Plot IDUs
    st.subheader('Resumo IDUs:')
    st.write('Quantidade total de IDUs por modelo.')
    sns.catplot(x="GAI", y="Count", data=resumo_idus, kind="bar", palette='deep',
                legend_out=False).set(title='IDUs').set_xticklabels(rotation=90)
    st.pyplot()
    if st.checkbox('Show IDUs Table'):
        st.write(resumo_idus)

    # Equipamento por Link
    st.subheader('Resumo Equipamento por Link:')
    st.write('Quantidade de IDUs, ODUs e Antenas para cada Link.')
    st.write(equipamento_link)
