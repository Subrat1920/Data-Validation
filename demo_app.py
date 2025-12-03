import streamlit as st
import pandas as pd

st.header("This is my personal Website")
st.write("Hi, welcome to the page")

file = st.file_uploader("Drop your CSV file here...", type=["csv"])

if file is not None:
    df = pd.read_csv(file)
    st.dataframe(df)
