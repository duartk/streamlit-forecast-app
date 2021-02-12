import streamlit as st
import pandas as pd
import numpy as np
import seaborn as sns
import base64
from io import BytesIO

def to_excel(df, df1, df2, df3):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Links')
    df1.to_excel(writer, sheet_name='Count_escopo')
    df2.to_excel(writer, sheet_name='Count_ant')
    df3.to_excel(writer, sheet_name='Count_conf')
    writer.save()
    processed_data = output.getvalue()
    return processed_data

def get_table_download_link(df, df1, df2, df3):
    """Generates a link allowing the data in a given panda dataframe to be downloaded
    in:  dataframe
    out: href string
    """
    val = to_excel(df, df1, df2, df3)
    b64 = base64.b64encode(val)  # val looks like b'...'
    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="extract.xlsx">Download file</a>' # decode b'abc' => abc
    # PS: pip install xlsxwriter  # pandas need this

# title for the app
st.title("Forecast App")

# Setup file upload
uploaded_file = st.sidebar.file_uploader(label='Upload your CSV or Excel file',
                                         type=['csv', 'xlsm', 'xlsx'])
st.sidebar.subheader("Visualization Settings")

global df
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        print(e)
        df = pd.read_csv(uploaded_file)

    columns_list = ['Regional', 'Link Type Engenharia', 'FREQUENCY', 'CONF', 'ANT TYPE', 'REUSO _ ANT A', 'ANT A',
                    'REUSO _ ANT B', 'ANT B', 'ANT SD A', 'ANT SD B']
    df1 = df[columns_list]

    antena_A = df1[df1['REUSO _ ANT A'] == "NO"][{'FREQUENCY', 'ANT A', 'ANT TYPE'}]
    antena_B = df1[df1['REUSO _ ANT B'] == "NO"][{'FREQUENCY', 'ANT B', 'ANT TYPE'}]
    antena_Ar = df1[df1['REUSO _ ANT A'] == "NO"][{'FREQUENCY', 'ANT SD A', 'ANT TYPE'}]
    antena_Br = df1[df1['REUSO _ ANT B'] == "NO"][{'FREQUENCY', 'ANT SD B', 'ANT TYPE'}]
    antena_A = antena_A.rename(columns={'ANT A': 'ANT', })
    antena_B = antena_B.rename(columns={'ANT B': 'ANT', })
    antena_Ar = antena_Ar.rename(columns={'ANT SD A': 'ANT', })
    antena_Br = antena_Br.rename(columns={'ANT SD B': 'ANT', })
    antenas = pd.concat([antena_A, antena_B, antena_Ar, antena_Br])
    antenas.dropna(inplace=True)
    antenas['FREQUENCY'].replace({"7.5": "7", "8": "7", "8.5": "7", "7/8": "7", "7/8.5": "7", "8/8.5": "7"},
                                 inplace=True)
    count_ant = antenas.groupby(['FREQUENCY', 'ANT TYPE'])['ANT'].value_counts(sort=False)
    count_conf = df1.groupby(['FREQUENCY'])['CONF'].value_counts(sort=False)
    cont_escopo = df1.groupby(['Regional'])['Link Type Engenharia'].value_counts()

    st.markdown(get_table_download_link(df, cont_escopo, count_ant, count_conf), unsafe_allow_html=True)

    sns.set_theme(style="darkgrid")
    g = sns.catplot(x="FREQUENCY", hue="ANT", data=antenas,
                    kind="count")
    # g = sns.catplot(x="ANT", col="FREQUENCY", data=antenas, col_wrap=3,
    #                 kind="count", height=5, aspect=.8)
    st.pyplot(g)

    option = st.sidebar.selectbox(
        'Select Frequency.',
        antenas['FREQUENCY'].unique())

    filtered_antenas = antenas[antenas['FREQUENCY'] == option]

    'You selected: ', option, 'Ghz'
    st.bar_chart(filtered_antenas['ANT'].value_counts())

    if st.sidebar.checkbox('Show raw data'):
        st.write(df1)
