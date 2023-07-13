import pandas as pd
import plotly.express as px  # pip install plotly-express
import streamlit as st  # pip install streamlit

st.set_page_config(page_title="Sales Dashboard",
                   page_icon=":bar_chart:",
                   layout="wide"
                   )

# p = pd.read_excel("D:/WorkSpace/R/Test/SalesDashboard/sales.xlsx")
# print(p)

# define the function to read data from the Excel file
@st.cache_data
def get_data_from_excel():
    df = pd.read_excel(
        io='D:/WorkSpace/R/Test/SalesDashboard/sales.xlsx',
        engine='openpyxl',
        skiprows=0,
        usecols='A:R',
        nrows=1500,
    )
    return df

# call the function to read the data from the Excel file and assign it to the variable df

df = get_data_from_excel()

st.sidebar.header("Please Filter Here:")
COUNTRY = st.sidebar.multiselect(

    "Select the Country:",
    options=df["COUNTRY"].unique(),
    default=df["COUNTRY"].unique()
)

STATE = st.sidebar.multiselect(
    "Select the State:",
    options=df["STATE"].unique(),
    default=df["STATE"].unique()
)

DEALSIZE = st.sidebar.multiselect(
    "Select the Dealsize:",
    options=df["DEALSIZE"].unique(),
    default=df["DEALSIZE"].unique()
)

df_selection = df.query(
    "COUNTRY == @COUNTRY & STATE == @STATE & DEALSIZE == @DEALSIZE"
)

st.dataframe(df_selection)

# --- MAINPAGE ----------
st.title(":bar_chart: Sales Dashboard")
st.markdown("##")

#  TOP KPI's
total_sales = int(df_selection["SALES"].sum())

# df['SALES'].astype(int)
# Group the DataFrame by country column and calculate the sum of sales
grouped = df_selection.groupby('COUNTRY')['SALES'].sum().rename("country_sales")

#  Group the DataFrame by country & state column and calculate the sum of sales
grouped2 = df_selection.groupby(['COUNTRY', 'STATE'])['SALES'].sum().rename("state_sales").reset_index()

grouped3 = df_selection.groupby(['COUNTRY', 'STATE', 'DEALSIZE'])['SALES'].sum().rename("dealsize_sales").reset_index()

# Set the index of the resulting Series using the index attribute
grouped_with_index = pd.Series(grouped.values, index=grouped.index)

grouped_with_index2 = grouped2.rename(columns={'SALES': 'state_sales'})

grouped_with_index3 = grouped3.rename(columns={'DEALSIZE': 'dealsize_sale'})

# Convert the resulting Series to a DataFrame with a new column name
df_sales_by_country = pd.DataFrame({'country_sales': grouped_with_index})

df_sales_by_state = pd.DataFrame(grouped_with_index2['state_sales'].values, index=grouped_with_index2.index, columns=['state_sales'])

df_sales_by_dealsize = pd.DataFrame(grouped_with_index3['dealsize_sale'].values, index=grouped_with_index3.index, columns=['dealsize_sale'])

# Display the DataFrame with the sum of country sales and the name of the country as the index
print(df_sales_by_country)

print(df_sales_by_state)

print(df_sales_by_dealsize)

left_column, middle_column, right_column, right2_column = st.columns(4)

with left_column:
    st.subheader("Total Sales:")
    st.subheader(f"US $ {total_sales:,}")

with middle_column:
    st.subheader("Country Sales:")
    st.subheader("Country Sales:")
    st.subheader(f"US $ {df_sales_by_country}")

with right_column:
    st.subheader("State Sales:")
    st.subheader("State Sales:")
    st.subheader(f"US $ {df_sales_by_state}")

with right2_column:
    st.subheader("Dealsize:")
    st.subheader("Dealsize:")
    st.subheader(f"US $ {df_sales_by_dealsize}")

st.markdown("---")

# SALES BY PRODUCTLINE [BAR CHART]
# df.groupby('PRODUCTLINE').sum()
# df.groupby('PRODUCTLINE').sum()[["SALES"]]
# df.groupby('PRODUCTLINE').sum()[["SALES"]].sort_values(by="SALES")

sales_by_product_line = (
    df_selection.groupby('PRODUCTLINE').sum().sort_values(by="SALES")
)

fig_product_sales = px.bar(
    sales_by_product_line,
    x="SALES",
    y=sales_by_product_line.index,
    orientation="h",
    title="<b>Sales by Product Line</b>",
    color_discrete_sequence=["#0083B8"] * len(sales_by_product_line),
    template="plotly_white",
)

fig_product_sales.update_layout(
    plot_bgcolor="rgba(0,0,0,0)",
    xaxis=(dict(showgrid=False))
)
st.plotly_chart(fig_product_sales)

# --- HIDE STREAMLIT STYLE ---
hide_st_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
"""
st.markdown(hide_st_style, unsafe_allow_html=True)
