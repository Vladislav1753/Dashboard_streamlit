import pandas as pd
import openpyxl
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Sales Dashboard",
                   page_icon=":bar_chart:",
                   layout="wide")

excel_path = r"D:\DS_DA_projects\Supermarket_sales_dashboard\supermarkt_sales.xlsx"

@st.cache
def get_data_from_excel():
    sales = pd.read_excel(io=excel_path,
                          engine='openpyxl',
                          sheet_name='Sales',
                          skiprows=3,
                          usecols="B:R",
                          nrows=1000)
    sales["Hour"] = pd.to_datetime(sales["Time"], format="%H:%M:%S").dt.hour
    return sales

sales = get_data_from_excel()
st.sidebar.header("Filters:")
city = st.sidebar.multiselect(
    "Select the city:",
    options=sales["City"].unique(),
    default=sales["City"].unique()
)

gender = st.sidebar.multiselect(
    "Select the gender:",
    options=sales["Gender"].unique(),
    default=sales["Gender"].unique()
)


customer_type = st.sidebar.multiselect(
    "Select the customer type:",
    options=sales["Customer_type"].unique(),
    default=sales["Customer_type"].unique()
)

sales_selection = sales.query(
    "City == @city & Customer_type == @customer_type & Gender == @gender"
)

st.title(":bar_chart: Sales Dashboard")
st.markdown("##")

total_sales = int(sales_selection['Total'].sum())
average_rating = round(sales_selection["Rating"].mean(), 1)
star_rating = ":star:" * int(round(average_rating, 0))
average_sale_by_transaction = round(sales_selection["Total"].mean(), 2)

left_column, middle_column, right_column = st.columns(3)
with left_column:
    st.subheader("Total Sales:")
    st.subheader(f"US $ {total_sales:,}")
with middle_column:
    st.subheader("Average Rating:")
    st.subheader(f"{average_rating} {star_rating}")
with right_column:
    st.subheader("Average Sales Per Transaction:")
    st.subheader(f"US $ {average_sale_by_transaction:,}")

st.markdown("---")

sales_by_product_line = (
    sales_selection.groupby(by=['Product line'])[["Total"]].sum().sort_values(by="Total")
)

fig_product_sales = px.bar(
    sales_by_product_line,
    x="Total",
    y=sales_by_product_line.index,
    orientation='h',
    title="<b>Sales by Product Line</b>",
    color_discrete_sequence=["#0083B8"] * len(sales_by_product_line),
    template="plotly_white"
)

fig_product_sales.update_layout(
    plot_bgcolor="rgba(0,0,0,0)",
    xaxis=(dict(showgrid=False))
)


sales_by_hour = sales_selection.groupby(by=["Hour"])[["Total"]].sum().sort_values(by=["Total"])
fig_hourly_sales = px.bar(
    sales_by_hour,
    x=sales_by_hour.index,
    y="Total",
    title="<b>Sales by hour</b>",
    color_discrete_sequence=["#0083B8"] * len(sales_by_hour),
    template="plotly_white"
)

fig_hourly_sales.update_layout(
    xaxis=dict(tickmode="linear"),
    plot_bgcolor="rgba(0,0,0,0)",
    yaxis=(dict(showgrid=False))
)


left_column, right_column = st.columns(2)
left_column.plotly_chart(fig_hourly_sales, use_container_width=True)
right_column.plotly_chart(fig_product_sales, use_container_width=True)