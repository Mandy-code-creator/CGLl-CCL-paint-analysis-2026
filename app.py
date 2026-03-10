```python
import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re

# ==========================================================
# PAGE
# ==========================================================

st.set_page_config(
    page_title="Length Variance Analysis",
    layout="wide"
)

st.title("Length Variance Analysis: Total CGL vs CCL per Order")

# ==========================================================
# GOOGLE SHEET
# ==========================================================

GSHEET_URL = "https://docs.google.com/spreadsheets/d/1-kayrLVYwOO66Xxc7Vk7dbTNZ5Aph4MVd9DMTz6RJS0/edit#gid=0"

# ==========================================================
# LOAD DATA
# ==========================================================

@st.cache_data(ttl=300)
def load_sheet(url):

    base = url.split("/edit")[0]
    csv = f"{base}/export?format=csv"

    df = pd.read_csv(csv)

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.lower()
        .str.replace(" ", "")
    )

    return df


df = load_sheet(GSHEET_URL)

# ==========================================================
# COLUMN FINDER
# ==========================================================

def find_col(names):

    for n in names:
        n = n.lower().replace(" ", "")
        if n in df.columns:
            return n

    return None


order_c = find_col(["訂單號碼","订单号码","order"])
mother_c = find_col(["投入鋼捲號碼","投入钢卷号码","mother"])
baby_c = find_col(["產出鋼捲號碼","产出钢卷号码","baby"])

cgl_l = find_col(["镀锌測長度","镀锌實測長度","鍍鋅測長度"])
ccl_l = find_col(["實測長度","实测长度"])

cgl_w = find_col(["镀锌測寬度","鍍鋅測寬度"])
cgl_t = find_col(["镀锌實測厚度","鍍鋅實測厚度"])

ccl_w = find_col(["實測寬度"])
ccl_t = find_col(["實測厚度"])

outer_cut = find_col(["outercutlength","outercut"])
inner_cut = find_col(["innercutlength","innercut"])

# ==========================================================
# NUMERIC
# ==========================================================

num_cols = [cgl_l,ccl_l,cgl_w,cgl_t,ccl_w,ccl_t,outer_cut,inner_cut]

for c in num_cols:
    if c:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

# ==========================================================
# BASIC CLEAN
# ==========================================================

df["base_coil"] = df[mother_c].astype(str).str[:-3]

df["is_x00"] = df[mother_c].astype(str).str.endswith("X00")

df["is_first_mother"] = ~df.duplicated([order_c,mother_c])

df["is_first_baby"] = ~df.duplicated([order_c,baby_c])

df.loc[~df["is_first_baby"],ccl_l] = 0

# ==========================================================
# FAMILY INFO
# ==========================================================

df["family_has_x00"] = df.groupby([order_c,"base_coil"])["is_x00"].transform("any")

df["family_has_cgl"] = df.groupby([order_c,"base_coil"])[cgl_l].transform(lambda x:(x>0).any())

df["sum_ccl_by_mother"] = df.groupby([order_c,mother_c])[ccl_l].transform("sum")

# ==========================================================
# INPUT RESOLVE
# ==========================================================

def resolve(row):

    if not row["is_first_mother"]:
        return 0

    if row[cgl_l] > 0:

        if row["family_has_x00"] and not row["is_x00"]:
            return 0

        return row[cgl_l]

    if row[cgl_l] == 0:

        if row["family_has_cgl"]:
            return 0

        return row["sum_ccl_by_mother"]

    return 0


df[cgl_l] = df.apply(resolve, axis=1)

# ==========================================================
# CUT FIRST ONLY
# ==========================================================

df.loc[~df["is_first_mother"],outer_cut] = 0
df.loc[~df["is_first_mother"],inner_cut] = 0

# ==========================================================
# MOTHER SUMMARY
# ==========================================================

mother = df.groupby([order_c,mother_c]).agg({

    cgl_l:"first",
    ccl_l:"sum",
    outer_cut:"max",
    inner_cut:"max",
    cgl_w:"mean",
    cgl_t:"mean",
    ccl_t:"mean"

}).reset_index()

# ==========================================================
# ORDER SUMMARY
# ==========================================================

summary = mother.groupby(order_c).agg({

    mother_c:"count",
    cgl_l:"sum",
    ccl_l:"sum",
    outer_cut:"sum",
    inner_cut:"sum",
    cgl_w:"mean",
    cgl_t:"mean",
    ccl_t:"mean"

}).reset_index()

summary = summary.rename(columns={

    mother_c:"Qty",
    cgl_l:"Input",
    ccl_l:"Output"

})

summary["Scrap"] = summary[outer_cut] + summary[inner_cut]

summary["Diff"] = summary["Output"] - (summary["Input"] - summary["Scrap"])

summary["Thick Var"] = summary[ccl_t] - summary[cgl_t]

summary["Area Diff"] = (summary[cgl_w]/1000) * summary["Diff"]

# ==========================================================
# DISPLAY
# ==========================================================

st.subheader("Order Summary")

disp = summary[[

    order_c,
    "Qty",
    cgl_w,
    "Input",
    "Scrap",
    "Output",
    "Diff",
    "Thick Var",
    "Area Diff"

]]

disp.columns = [

    "Order ID",
    "Qty(Coils)",
    "Width(mm)",
    "Input(m)",
    "Scrap(m)",
    "Output(m)",
    "Diff(m)",
    "Thick Var",
    "Diff Area(m²)"

]

disp = disp.sort_values("Scrap(m)",ascending=False)

st.dataframe(disp,use_container_width=True)

# ==========================================================
# KPI
# ==========================================================

st.divider()

c1,c2,c3,c4 = st.columns(4)

c1.metric("Orders",len(disp))
c2.metric("Total Input",f"{disp['Input(m)'].sum():,.0f} m")
c3.metric("Total Output",f"{disp['Output(m)'].sum():,.0f} m")
c4.metric("Total Scrap",f"{disp['Scrap(m)'].sum():,.0f} m")

# ==========================================================
# CHART
# ==========================================================

st.divider()

st.subheader("Scrap by Order")

fig = px.bar(

    disp,
    x="Order ID",
    y="Scrap(m)",
    color="Diff(m)"

)

st.plotly_chart(fig,use_container_width=True)

# ==========================================================
# EXPORT
# ==========================================================

st.divider()

buffer = io.BytesIO()

with pd.ExcelWriter(buffer,engine="xlsxwriter") as writer:

    disp.to_excel(writer,index=False)

st.download_button(

    "Download Excel",
    buffer.getvalue(),
    "report.xlsx"

)
```
