import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import date
from io import BytesIO
import os

# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="BUSCH RDC Picking Efficiency Model",
    layout="wide"
)

PASSWORD = "Buschrdc@2026"

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False


# =========================================================
# HELPER FUNCTIONS
# =========================================================

def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Report")
    return output.getvalue()


def convert_multiple_to_excel(sheet_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheet_dict.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    return output.getvalue()


def clean_storage_type(value):
    value = str(value).strip()

    if value in ["0100", "100", "Pallet Rack", "PALLET RACK"]:
        return "Pallet Rack"

    if value in ["0200", "200", "Small Bin", "SMALL BIN"]:
        return "Small Bin"

    return value


def clean_master_value(value):
    if pd.isna(value):
        return "Blank"

    value = str(value).strip()

    if value.lower() in ["nan", "none", "null", ""]:
        return "Blank"

    if value.endswith(".0"):
        value = value[:-2]

    return value


def get_reference_file_path():
    current_folder_path = os.path.join(
        os.path.dirname(__file__),
        "Reference_Data",
        "RDC_Master_Reference.xlsx"
    )

    parent_folder_path = os.path.join(
        os.path.dirname(os.path.dirname(__file__)),
        "Reference_Data",
        "RDC_Master_Reference.xlsx"
    )

    if os.path.exists(current_folder_path):
        return current_folder_path

    return parent_folder_path


# =========================================================
# LOGIN
# =========================================================

if not st.session_state.authenticated:

    st.markdown(
        """
        <div style='background-color:#ff6600; padding:35px; border-radius:10px; text-align:center;'>
            <h1 style='color:white; font-size:52px; font-weight:900; margin-bottom:0px;'>
                BUSCH GROUP REGIONAL DISTRIBUTION CENTER
            </h1>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.title("Dynamic Slotting, Picking Efficiency & Warehouse Model")

    password = st.text_input("Enter Password", type="password")

    if st.button("Login"):
        if password == PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Incorrect Password")


# =========================================================
# MAIN APP
# =========================================================

else:

    st.markdown(
        """
        <div style='background-color:#ff6600; padding:35px; border-radius:10px; text-align:center;'>
            <h1 style='color:white; font-size:54px; font-weight:1000; margin-bottom:0px;'>
                BUSCH GROUP REGIONAL DISTRIBUTION CENTER
            </h1>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.title("Dynamic Slotting, Picking Efficiency & Warehouse Model")

    tab1, tab2, tab3, tab4 = st.tabs(
        [
            "Picking Efficiency Model",
            "Capacity Modeling",
            "About RDC",
            "Historical Runs"
        ]
    )

    # =====================================================
    # TAB 1 — PICKING EFFICIENCY MODEL
    # =====================================================

    with tab1:

        st.header("Analysis Setup")

        col1, col2, col3 = st.columns(3)

        with col1:
            from_date = st.date_input("From Date", value=date(2025, 10, 1))

        with col2:
            to_date = st.date_input("To Date", value=date.today())

        with col3:
            save_name = st.text_input(
                "Save Analysis As",
                value=f"{from_date}_to_{to_date}_Data"
            )

        uploaded_file = st.file_uploader(
            "Upload RDC Input Template",
            type=["xlsx"]
        )

        if uploaded_file is not None:

            product_locations = pd.read_excel(uploaded_file, sheet_name="Product_Locations")
            product_picks = pd.read_excel(uploaded_file, sheet_name="Product_Picks")
            bin_master = pd.read_excel(uploaded_file, sheet_name="Bin_Master")
            bin_picks = pd.read_excel(uploaded_file, sheet_name="Bin_Picks")
            product_master = pd.read_excel(uploaded_file, sheet_name="Product_Master")

            product_locations.columns = product_locations.columns.str.strip()
            product_picks.columns = product_picks.columns.str.strip()
            bin_master.columns = bin_master.columns.str.strip()
            bin_picks.columns = bin_picks.columns.str.strip()
            product_master.columns = product_master.columns.str.strip()

            product_locations["Product"] = product_locations["Product"].astype(str).str.strip()
            product_locations["Bin Location"] = product_locations["Bin Location"].astype(str).str.strip()
            product_picks["Product"] = product_picks["Product"].astype(str).str.strip()
            bin_master["Bin Location"] = bin_master["Bin Location"].astype(str).str.strip()
            bin_picks["Bin Location"] = bin_picks["Bin Location"].astype(str).str.strip()

            product_locations["Storage Type"] = product_locations["Storage Type"].apply(clean_storage_type)
            bin_master["Storage Type"] = bin_master["Storage Type"].apply(clean_storage_type)
            bin_picks["Storage Type"] = bin_picks["Storage Type"].apply(clean_storage_type)

            product_picks["Total Picks"] = pd.to_numeric(
                product_picks["Total Picks"],
                errors="coerce"
            ).fillna(0)

            if "Access Classification" in bin_master.columns:
                bin_master["Access Classification"] = bin_master["Access Classification"].apply(clean_master_value)
            else:
                bin_master["Access Classification"] = "Missing Bin Master"

            if "Blocked Bin" in bin_master.columns:
                bin_master["Blocked Bin"] = bin_master["Blocked Bin"].apply(clean_master_value)
            else:
                bin_master["Blocked Bin"] = "No"

            # =================================================
            # KPI LOGIC
            # =================================================

            unique_rdc_products = product_locations["Product"].nunique()

            picked_product_summary = (
                product_picks
                .groupby("Product", as_index=False)["Total Picks"]
                .sum()
            )

            products_with_picks = picked_product_summary[
                picked_product_summary["Total Picks"] > 0
            ]["Product"].nunique()

            products_without_picks = unique_rdc_products - products_with_picks

            total_product_picks = picked_product_summary["Total Picks"].sum()

            total_days = (to_date - from_date).days + 1
            total_weeks = total_days / 7
            total_months = total_days / 30.44

            avg_day = round(total_product_picks / total_days, 1)
            avg_week = round(total_product_picks / total_weeks, 1)
            avg_month = round(total_product_picks / total_months, 1)

            pallet_locations = bin_master[
                bin_master["Storage Type"] == "Pallet Rack"
            ]["Bin Location"].nunique()

            small_bin_levels = bin_master[
                bin_master["Storage Type"] == "Small Bin"
            ]["Bin Location"].nunique()

            st.markdown("---")
            st.header("Executive Warehouse KPI Dashboard")

            k1, k2, k3, k4 = st.columns(4)

            k1.metric("Unique Products in RDC Locations", f"{unique_rdc_products:,}")
            k2.metric("Products With Picks", f"{products_with_picks:,}")
            k3.metric("Products Without Picks", f"{products_without_picks:,}")
            k4.metric("Total Product Picks", f"{int(total_product_picks):,}")

            k5, k6, k7 = st.columns(3)

            k5.metric("Available Pallet Rack Locations", f"{pallet_locations:,}")
            k6.metric("Available Small Bin Levels", f"{small_bin_levels:,}")
            k7.metric("Analysis Period Days", f"{total_days:,}")

            st.markdown("---")
            st.header("Pick Rate Intelligence")

            p1, p2, p3 = st.columns(3)

            p1.metric("Average Picks per Day", f"{avg_day:,}")
            p2.metric("Average Picks per Week", f"{avg_week:,}")
            p3.metric("Average Picks per Month", f"{avg_month:,}")

            # =================================================
            # ABCD ANALYSIS
            # =================================================

            picked_only = picked_product_summary[
                picked_product_summary["Total Picks"] > 0
            ].copy()

            d_items = picked_only[
                picked_only["Total Picks"] == 1
            ].copy()

            d_items["ABC Classification"] = "D - One Pick Item"

            abc_base = picked_only[
                picked_only["Total Picks"] > 1
            ].copy()

            abc_base = abc_base.sort_values("Total Picks", ascending=False).reset_index(drop=True)

            abc_volume_total = abc_base["Total Picks"].sum()

            abc_base["Cumulative Picks"] = abc_base["Total Picks"].cumsum()
            abc_base["Cumulative Pick %"] = abc_base["Cumulative Picks"] / abc_volume_total

            def classify_abc(row):
                if row["Cumulative Pick %"] <= 0.80:
                    return "A - Top 80% Pick Volume"
                elif row["Cumulative Pick %"] <= 0.95:
                    return "B - Next 15% Pick Volume"
                else:
                    return "C - Remaining Pick Volume"

            abc_base["ABC Classification"] = abc_base.apply(classify_abc, axis=1)

            d_items["Cumulative Picks"] = 0
            d_items["Cumulative Pick %"] = 0

            abc_final = pd.concat([abc_base, d_items], ignore_index=True)

            abc_summary = (
                abc_final
                .groupby("ABC Classification")
                .agg(
                    Product_Count=("Product", "nunique"),
                    Total_Picks=("Total Picks", "sum")
                )
                .reset_index()
            )

            abc_categories = pd.DataFrame({
                "ABC Classification": [
                    "A - Top 80% Pick Volume",
                    "B - Next 15% Pick Volume",
                    "C - Remaining Pick Volume",
                    "D - One Pick Item"
                ]
            })

            abc_summary = abc_categories.merge(
                abc_summary,
                on="ABC Classification",
                how="left"
            ).fillna(0)

            abc_summary["Product_Count"] = abc_summary["Product_Count"].astype(int)
            abc_summary["Total_Picks"] = abc_summary["Total_Picks"].astype(int)

            st.markdown("---")
            st.header("ABCD Product Velocity Analysis")

            st.write("ABCD analysis is calculated only from the Product_Picks tab.")

            c1, c2 = st.columns(2)

            with c1:
                fig_abc_bar = px.bar(
                    abc_summary,
                    x="ABC Classification",
                    y="Product_Count",
                    text="Product_Count",
                    title="Product Count by ABCD Classification"
                )
                fig_abc_bar.update_traces(texttemplate="%{text:.0f}", textposition="outside")
                st.plotly_chart(fig_abc_bar, use_container_width=True)

            with c2:
                fig_abc_pie = px.pie(
                    abc_summary,
                    names="ABC Classification",
                    values="Total_Picks",
                    title="Pick Volume Share by ABCD Classification"
                )
                st.plotly_chart(fig_abc_pie, use_container_width=True)

            fig_abc_pick_bar = px.bar(
                abc_summary,
                x="ABC Classification",
                y="Total_Picks",
                text="Total_Picks",
                title="Number of Lines Picked by ABCD Classification"
            )
            fig_abc_pick_bar.update_traces(texttemplate="%{text:.0f}", textposition="outside")
            st.plotly_chart(fig_abc_pick_bar, use_container_width=True)

            st.dataframe(abc_summary, use_container_width=True, hide_index=True)

            st.success(
                f"ABCD pick volume check: {int(abc_summary['Total_Picks'].sum()):,} total picks classified. "
                f"This matches Product_Picks total: {int(total_product_picks):,}."
            )

            # =================================================
            # MATERIAL MASTER ANALYSIS
            # =================================================

            st.markdown("---")
            st.header("Material Master Analysis")

            for col in [
                "Plant-Sp.Matl Status",
                "X-Plant Material Status",
                "Procurement Type",
                "Country of origin"
            ]:
                if col in product_master.columns:
                    product_master[col] = product_master[col].apply(clean_master_value)
                else:
                    product_master[col] = "Blank"

            plant_status_summary = (
                product_master
                .groupby("Plant-Sp.Matl Status", dropna=False)
                .size()
                .reset_index(name="Material Count")
                .sort_values("Material Count", ascending=False)
            )

            global_status_summary = (
                product_master
                .groupby("X-Plant Material Status", dropna=False)
                .size()
                .reset_index(name="Material Count")
                .sort_values("Material Count", ascending=False)
            )

            procurement_summary = (
                product_master
                .groupby("Procurement Type", dropna=False)
                .size()
                .reset_index(name="Material Count")
                .sort_values("Material Count", ascending=False)
            )

            country_summary = (
                product_master
                .groupby("Country of origin", dropna=False)
                .size()
                .reset_index(name="Material Count")
                .sort_values("Material Count", ascending=False)
            )

            m1, m2 = st.columns(2)

            with m1:
                st.subheader("Plant-Specific Material Status")
                st.dataframe(plant_status_summary, use_container_width=True, hide_index=True)

            with m2:
                st.subheader("Global Material Status")
                st.dataframe(global_status_summary, use_container_width=True, hide_index=True)

            m3, m4 = st.columns(2)

            with m3:
                st.subheader("Procurement Type")
                st.dataframe(procurement_summary, use_container_width=True, hide_index=True)

            with m4:
                st.subheader("Country of Origin")
                st.dataframe(country_summary, use_container_width=True, hide_index=True)

            # =================================================
            # ACCESS CLASSIFICATION
            # =================================================

            product_locations = product_locations.merge(
                bin_master[
                    [
                        "Bin Location",
                        "Access Classification",
                        "Blocked Bin"
                    ]
                ],
                on="Bin Location",
                how="left"
            )

            product_locations["Access Classification"] = product_locations["Access Classification"].fillna("Missing Bin Master")
            product_locations["Blocked Bin"] = product_locations["Blocked Bin"].fillna("Blank")

            product_locations["Access Classification"] = product_locations.apply(
                lambda row: "Blocked"
                if str(row["Blocked Bin"]).strip().lower() == "yes"
                else row["Access Classification"],
                axis=1
            )

            # =================================================
            # STORAGE CATEGORY
            # =================================================

            product_storage = (
                product_locations
                .groupby("Product")["Storage Type"]
                .apply(lambda x: sorted(set(x)))
                .reset_index()
            )

            def storage_category(storage_list):
                if storage_list == ["Small Bin"]:
                    return "Exclusive Small Bin Product"
                elif storage_list == ["Pallet Rack"]:
                    return "Exclusive Pallet Rack Product"
                elif "Small Bin" in storage_list and "Pallet Rack" in storage_list:
                    return "Mixed Storage Product"
                else:
                    return "Review Storage Type"

            product_storage["Storage Category"] = product_storage["Storage Type"].apply(storage_category)

            storage_summary = (
                product_storage
                .groupby("Storage Category")
                .size()
                .reset_index(name="Product Count")
                .sort_values("Product Count", ascending=False)
            )

            st.markdown("---")
            st.header("Small Bin vs Pallet Rack Product Logic")
            st.dataframe(storage_summary, use_container_width=True, hide_index=True)

            # =================================================
            # SLOT DATA
            # =================================================

            merged = product_locations.merge(
                abc_final[["Product", "Total Picks", "ABC Classification"]],
                on="Product",
                how="left"
            )

            merged["Total Picks"] = merged["Total Picks"].fillna(0)
            merged["ABC Classification"] = merged["ABC Classification"].fillna("No Picks")

            merged = merged.merge(
                product_storage[["Product", "Storage Category"]],
                on="Product",
                how="left"
            )

            bad_access_types = [
                "Restricted Ground Access",
                "Low Accessibility Reserve",
                "Extended Small Bin Access",
                "Limited Efficiency Access",
                "Extended Reserve Access",
                "Blocked"
            ]

            prime_access_types = [
                "Prime Ground Access",
                "Prime Small Bin Access"
            ]

            def priority_logic(row):

                abc = row["ABC Classification"]
                access = row["Access Classification"]

                if abc == "A - Top 80% Pick Volume" and access in bad_access_types:
                    return 1

                if abc == "A - Top 80% Pick Volume" and access not in prime_access_types:
                    return 2

                if abc == "B - Next 15% Pick Volume" and access in bad_access_types:
                    return 3

                if row["Storage Category"] == "Mixed Storage Product":
                    return 4

                if abc in ["D - One Pick Item", "No Picks"] and access in prime_access_types:
                    return 5

                return 99

            def warehouse_action(priority):

                if priority == 1:
                    return "Fix first: A item in bad access"
                elif priority == 2:
                    return "Move A item closer to prime access"
                elif priority == 3:
                    return "Move B item to better access"
                elif priority == 4:
                    return "Review mixed small bin / pallet rack storage"
                elif priority == 5:
                    return "Move low/no-pick item out of good access"
                else:
                    return "Acceptable / Monitor"

            def recommended_future_access(row):

                abc = row["ABC Classification"]
                storage = row["Storage Type"]
                storage_category_value = row["Storage Category"]

                if storage_category_value == "Mixed Storage Product":
                    if storage == "Small Bin":
                        return "Review qty: keep in Prime / Standard Small Bin or move to Pallet Rack"
                    else:
                        return "Review qty: keep in Prime / Standard Ground or move to Small Bin"

                if abc == "A - Top 80% Pick Volume":
                    if storage == "Small Bin":
                        return "Prime Small Bin Access"
                    else:
                        return "Prime Ground Access"

                if abc == "B - Next 15% Pick Volume":
                    if storage == "Small Bin":
                        return "Standard Small Bin Access"
                    else:
                        return "Standard Ground Access / High Efficiency Access"

                if abc == "C - Remaining Pick Volume":
                    if storage == "Small Bin":
                        return "Standard Small Bin Access / Extended Small Bin Access"
                    else:
                        return "Balanced Efficiency Access / Managed Reserve Access"

                if abc == "D - One Pick Item":
                    if storage == "Small Bin":
                        return "Extended Small Bin Access"
                    else:
                        return "Extended Reserve Access / Low Accessibility Reserve"

                if abc == "No Picks":
                    if storage == "Small Bin":
                        return "Extended Small Bin Access"
                    else:
                        return "Low Accessibility Reserve"

                return "Review manually"

            merged["Priority Group"] = merged.apply(priority_logic, axis=1)
            merged["Warehouse Action"] = merged["Priority Group"].apply(warehouse_action)
            merged["Recommended Future Access Classification"] = merged.apply(recommended_future_access, axis=1)

            action_list = (
                merged[merged["Priority Group"] != 99]
                .sort_values(
                    by=["Priority Group", "Total Picks"],
                    ascending=[True, False]
                )
                .reset_index(drop=True)
            )

            action_list["Action Sequence"] = range(1, len(action_list) + 1)

            priority_summary = (
                action_list
                .groupby(["Priority Group", "Warehouse Action"])
                .agg(
                    Product_Count=("Product", "nunique"),
                    Location_Count=("Bin Location", "nunique")
                )
                .reset_index()
                .sort_values("Priority Group")
            )

            st.markdown("---")
            st.header("Warehouse Improvement Workload")

            fig_priority = px.bar(
                priority_summary,
                x="Priority Group",
                y="Product_Count",
                text="Product_Count",
                title="Products Requiring Action by Priority"
            )
            fig_priority.update_traces(texttemplate="%{text:.0f}", textposition="outside")
            st.plotly_chart(fig_priority, use_container_width=True)

            st.dataframe(priority_summary, use_container_width=True, hide_index=True)

            # =================================================
            # BIN UTILIZATION BY ACCESS TYPE
            # =================================================

            expected_access = pd.DataFrame({
                "Access Classification": [
                    "Prime Ground Access",
                    "Standard Ground Access",
                    "Restricted Ground Access",
                    "High Efficiency Access",
                    "Balanced Efficiency Access",
                    "Limited Efficiency Access",
                    "Managed Reserve Access",
                    "Extended Reserve Access",
                    "Low Accessibility Reserve",
                    "Prime Small Bin Access",
                    "Standard Small Bin Access",
                    "Extended Small Bin Access",
                    "Blocked"
                ]
            })

            occupied_bins = product_locations[["Bin Location"]].drop_duplicates()

            bin_utilization = bin_master.merge(
                occupied_bins.assign(Used=1),
                on="Bin Location",
                how="left"
            )

            bin_utilization["Used"] = bin_utilization["Used"].fillna(0)

            bin_utilization["Access Classification"] = bin_utilization.apply(
                lambda row: "Blocked"
                if str(row["Blocked Bin"]).strip().lower() == "yes"
                else row["Access Classification"],
                axis=1
            )

            access_utilization = (
                bin_utilization
                .groupby("Access Classification")
                .agg(
                    Total_Bins=("Bin Location", "nunique"),
                    Used_Bins=("Used", "sum")
                )
                .reset_index()
            )

            access_utilization = expected_access.merge(
                access_utilization,
                on="Access Classification",
                how="left"
            ).fillna(0)

            access_utilization["Total_Bins"] = access_utilization["Total_Bins"].astype(int)
            access_utilization["Used_Bins"] = access_utilization["Used_Bins"].astype(int)

            access_utilization["Utilization %"] = np.where(
                access_utilization["Total_Bins"] == 0,
                0,
                (
                    access_utilization["Used_Bins"] /
                    access_utilization["Total_Bins"] * 100
                ).round(1)
            )

            st.markdown("---")
            st.header("Bin Utilization by Access Type")

            fig_util = px.bar(
                access_utilization,
                x="Utilization %",
                y="Access Classification",
                orientation="h",
                text="Utilization %",
                title="Bin Utilization by Access Type"
            )

            fig_util.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
            st.plotly_chart(fig_util, use_container_width=True)

            st.dataframe(access_utilization, use_container_width=True, hide_index=True)

            # =================================================
            # ACTION LIST
            # =================================================

            st.markdown("---")
            st.header("Warehouse Action Priority List")

            with st.expander("Filters for Warehouse Action Priority List"):

                priority_filter = st.multiselect(
                    "Priority Group",
                    options=sorted(action_list["Priority Group"].unique()),
                    default=sorted(action_list["Priority Group"].unique())
                )

                abc_filter = st.multiselect(
                    "ABC Classification",
                    options=sorted(action_list["ABC Classification"].unique()),
                    default=sorted(action_list["ABC Classification"].unique())
                )

                storage_filter = st.multiselect(
                    "Storage Type",
                    options=sorted(action_list["Storage Type"].unique()),
                    default=sorted(action_list["Storage Type"].unique())
                )

            filtered_action_list = action_list[
                (action_list["Priority Group"].isin(priority_filter)) &
                (action_list["ABC Classification"].isin(abc_filter)) &
                (action_list["Storage Type"].isin(storage_filter))
            ]

            final_action_columns = [
                "Action Sequence",
                "Priority Group",
                "Product",
                "Bin Location",
                "Storage Type",
                "Access Classification",
                "Recommended Future Access Classification",
                "Total Picks",
                "ABC Classification",
                "Storage Category",
                "Warehouse Action"
            ]

            filtered_action_list = filtered_action_list[final_action_columns]

            st.dataframe(
                filtered_action_list,
                use_container_width=True,
                height=500,
                hide_index=True
            )

            # =================================================
            # CONSOLIDATION
            # =================================================

            consolidation = (
                product_locations
                .groupby("Product")
                .agg(
                    Bin_Count=("Bin Location", "nunique"),
                    Storage_Types=("Storage Type", lambda x: ", ".join(sorted(set(x))))
                )
                .reset_index()
            )

            consolidation = consolidation.merge(
                abc_final[["Product", "Total Picks", "ABC Classification"]],
                on="Product",
                how="left"
            )

            consolidation["Total Picks"] = consolidation["Total Picks"].fillna(0)
            consolidation["ABC Classification"] = consolidation["ABC Classification"].fillna("No Picks")

            consolidation = consolidation.merge(
                product_storage[["Product", "Storage Category"]],
                on="Product",
                how="left"
            )

            consolidation = consolidation[
                consolidation["Bin_Count"] > 1
            ].copy()

            consolidation["Consolidation Recommendation"] = (
                "If quantity allows, consolidate into fewer bin locations"
            )

            consolidation = consolidation.sort_values(
                by=["Total Picks", "Bin_Count"],
                ascending=[False, False]
            ).reset_index(drop=True)

            consolidation["Consolidation Priority"] = range(1, len(consolidation) + 1)

            st.markdown("---")
            st.header("Consolidation Analysis")

            with st.expander("Filters for Consolidation Analysis"):

                consolidation_abc_filter = st.multiselect(
                    "Consolidation ABC Filter",
                    options=sorted(consolidation["ABC Classification"].unique()),
                    default=sorted(consolidation["ABC Classification"].unique())
                )

            filtered_consolidation = consolidation[
                consolidation["ABC Classification"].isin(consolidation_abc_filter)
            ]

            st.dataframe(
                filtered_consolidation,
                use_container_width=True,
                height=400,
                hide_index=True
            )

            # =================================================
            # DOWNLOADS
            # =================================================

            st.markdown("---")
            st.header("Download Reports")

            action_file = convert_df_to_excel(filtered_action_list)
            consolidation_file = convert_df_to_excel(filtered_consolidation)

            executive_kpis = pd.DataFrame(
                {
                    "Metric": [
                        "Unique Products in RDC Locations",
                        "Products With Picks",
                        "Products Without Picks",
                        "Total Product Picks",
                        "Average Picks per Day",
                        "Average Picks per Week",
                        "Average Picks per Month",
                        "Available Pallet Rack Locations",
                        "Available Small Bin Levels"
                    ],
                    "Value": [
                        unique_rdc_products,
                        products_with_picks,
                        products_without_picks,
                        int(total_product_picks),
                        avg_day,
                        avg_week,
                        avg_month,
                        pallet_locations,
                        small_bin_levels
                    ]
                }
            )

            leadership_file = convert_multiple_to_excel(
                {
                    "Executive KPIs": executive_kpis,
                    "ABCD Summary": abc_summary,
                    "Priority Summary": priority_summary,
                    "Bin Utilization": access_utilization
                }
            )

            full_file = convert_multiple_to_excel(
                {
                    "Executive KPIs": executive_kpis,
                    "ABCD Summary": abc_summary,
                    "Warehouse Action List": filtered_action_list,
                    "Consolidation": filtered_consolidation,
                    "Bin Utilization": access_utilization,
                    "Plant Status": plant_status_summary,
                    "Global Status": global_status_summary,
                    "Procurement Type": procurement_summary,
                    "Country Origin": country_summary,
                    "Product Picks": product_picks,
                    "Product Locations": product_locations,
                    "Bin Picks": bin_picks
                }
            )

            d1, d2, d3, d4 = st.columns(4)

            with d1:
                st.download_button(
                    "Download Action Priority List",
                    action_file,
                    "Warehouse_Action_Priority_List.xlsx"
                )

            with d2:
                st.download_button(
                    "Download Consolidation Analysis",
                    consolidation_file,
                    "Consolidation_Analysis.xlsx"
                )

            with d3:
                st.download_button(
                    "Download Leadership KPI Report",
                    leadership_file,
                    "Leadership_KPI_Report.xlsx"
                )

            with d4:
                st.download_button(
                    "Download Full Report",
                    full_file,
                    "RDC_Picking_Efficiency_Full_Report.xlsx"
                )

            # =================================================
            # SAVE HISTORICAL RUN
            # =================================================

            st.markdown("---")
            st.header("Save Historical Run")

            if st.button("Save Current Analysis Run"):

                history_folder = "Historical_Runs"
                os.makedirs(history_folder, exist_ok=True)

                run_folder = os.path.join(
                    history_folder,
                    save_name.replace(" ", "_")
                )

                os.makedirs(run_folder, exist_ok=True)

                with open(os.path.join(run_folder, "Leadership_KPI_Report.xlsx"), "wb") as f:
                    f.write(leadership_file)

                with open(os.path.join(run_folder, "Warehouse_Action_Priority_List.xlsx"), "wb") as f:
                    f.write(action_file)

                with open(os.path.join(run_folder, "Consolidation_Analysis.xlsx"), "wb") as f:
                    f.write(consolidation_file)

                with open(os.path.join(run_folder, "Full_Report.xlsx"), "wb") as f:
                    f.write(full_file)

                st.success(f"Saved historical run: {run_folder}")

        else:
            st.info("Please upload the RDC input template to continue.")

    # =====================================================
    # TAB 2 — CAPACITY MODELING
    # =====================================================

    with tab2:

        st.header("Capacity Modeling")

        st.write(
            "Estimate resource capacity by workstream: Receiving/Putaway, Pfeiffer Orders, Busch Orders, and Key Account Orders."
        )

        st.markdown("---")

        with st.form("capacity_model_form"):

            c1, c2, c3, c4 = st.columns(4)

            with c1:
                st.subheader("Receiving / Putaway")
                receiving_lines = st.number_input("Receiving Lines", min_value=0, value=70)
                receiving_minutes = st.number_input("Minutes per Receiving Line", min_value=0.1, value=5.0)
                putaway_lines = st.number_input("Putaway Lines", min_value=0, value=70)
                putaway_minutes = st.number_input("Minutes per Putaway Line", min_value=0.1, value=3.0)

            with c2:
                st.subheader("Pfeiffer Orders")
                pf_release_lines = st.number_input("Pfeiffer Release Lines", min_value=0, value=100)
                pf_release_minutes = st.number_input("Minutes per Pfeiffer Release Line", min_value=0.1, value=1.0)
                pf_pick_lines = st.number_input("Pfeiffer Pick Lines", min_value=0, value=100)
                pf_pick_minutes = st.number_input("Minutes per Pfeiffer Pick Line", min_value=0.1, value=5.0)
                pf_qa_lines = st.number_input("Pfeiffer QA Lines", min_value=0, value=100)
                pf_qa_minutes = st.number_input("Minutes per Pfeiffer QA Line", min_value=0.1, value=1.0)
                pf_pack_lines = st.number_input("Pfeiffer Pack Lines", min_value=0, value=100)
                pf_pack_minutes = st.number_input("Minutes per Pfeiffer Pack Line", min_value=0.1, value=3.0)
                pf_ship_lines = st.number_input("Pfeiffer Ship Lines", min_value=0, value=100)
                pf_ship_minutes = st.number_input("Minutes per Pfeiffer Ship Line", min_value=0.1, value=3.0)

            with c3:
                st.subheader("Busch Orders")
                busch_release_lines = st.number_input("Busch Release Lines", min_value=0, value=100)
                busch_release_minutes = st.number_input("Minutes per Busch Release Line", min_value=0.1, value=1.0)
                busch_pick_lines = st.number_input("Busch Pick Lines", min_value=0, value=100)
                busch_pick_minutes = st.number_input("Minutes per Busch Pick Line", min_value=0.1, value=5.0)
                busch_qa_lines = st.number_input("Busch QA Lines", min_value=0, value=100)
                busch_qa_minutes = st.number_input("Minutes per Busch QA Line", min_value=0.1, value=1.0)
                busch_pack_lines = st.number_input("Busch Pack Lines", min_value=0, value=100)
                busch_pack_minutes = st.number_input("Minutes per Busch Pack Line", min_value=0.1, value=3.0)
                busch_ship_lines = st.number_input("Busch Ship Lines", min_value=0, value=100)
                busch_ship_minutes = st.number_input("Minutes per Busch Ship Line", min_value=0.1, value=3.0)

            with c4:
                st.subheader("Key Account Orders")
                key_release_lines = st.number_input("Key Account Release Lines", min_value=0, value=100)
                key_release_minutes = st.number_input("Minutes per Key Account Release Line", min_value=0.1, value=1.0)
                key_pick_lines = st.number_input("Key Account Pick Lines", min_value=0, value=100)
                key_pick_minutes = st.number_input("Minutes per Key Account Pick Line", min_value=0.1, value=5.0)
                key_qa_lines = st.number_input("Key Account QA Lines", min_value=0, value=100)
                key_qa_minutes = st.number_input("Minutes per Key Account QA Line", min_value=0.1, value=1.0)
                key_pack_lines = st.number_input("Key Account Pack Lines", min_value=0, value=100)
                key_pack_minutes = st.number_input("Minutes per Key Account Pack Line", min_value=0.1, value=3.0)
                key_ship_lines = st.number_input("Key Account Ship Lines", min_value=0, value=100)
                key_ship_minutes = st.number_input("Minutes per Key Account Ship Line", min_value=0.1, value=3.0)

            st.markdown("---")

            a1, a2 = st.columns(2)

            with a1:
                available_hours_per_resource = st.number_input(
                    "Available Hours per Resource",
                    min_value=1.0,
                    value=7.0
                )

            with a2:
                current_resources = st.number_input(
                    "Current Available Resources",
                    min_value=0.0,
                    value=3.0
                )

            submitted = st.form_submit_button("Run Capacity Model")

        if submitted:

            receiving_hours = ((receiving_lines * receiving_minutes) + (putaway_lines * putaway_minutes)) / 60

            pfeiffer_hours = (
                (pf_release_lines * pf_release_minutes)
                + (pf_pick_lines * pf_pick_minutes)
                + (pf_qa_lines * pf_qa_minutes)
                + (pf_pack_lines * pf_pack_minutes)
                + (pf_ship_lines * pf_ship_minutes)
            ) / 60

            busch_hours = (
                (busch_release_lines * busch_release_minutes)
                + (busch_pick_lines * busch_pick_minutes)
                + (busch_qa_lines * busch_qa_minutes)
                + (busch_pack_lines * busch_pack_minutes)
                + (busch_ship_lines * busch_ship_minutes)
            ) / 60

            key_account_hours = (
                (key_release_lines * key_release_minutes)
                + (key_pick_lines * key_pick_minutes)
                + (key_qa_lines * key_qa_minutes)
                + (key_pack_lines * key_pack_minutes)
                + (key_ship_lines * key_ship_minutes)
            ) / 60

            total_resource_hours = receiving_hours + pfeiffer_hours + busch_hours + key_account_hours

            exact_resources = total_resource_hours / available_hours_per_resource
            planned_resources = int(np.ceil(exact_resources))

            available_resource_hours = current_resources * available_hours_per_resource
            resource_gap = total_resource_hours - available_resource_hours

            st.markdown("---")
            st.subheader("Capacity Results")

            r1, r2, r3, r4 = st.columns(4)

            r1.metric("Total Resource Hours Required", f"{total_resource_hours:,.1f}")
            r2.metric("Exact Resources Required", f"{exact_resources:,.1f}")
            r3.metric("Planned Resource Requirement", f"{planned_resources:,}")
            r4.metric("Resource Hour Gap / Surplus", f"{resource_gap:,.1f}")

            if resource_gap > 0:
                st.warning("Current resource capacity is short. Additional support or overtime may be required.")
            else:
                st.success("Current resource capacity is enough for the entered workload.")

            capacity_breakdown = pd.DataFrame({
                "Workstream": [
                    "Receiving / Putaway",
                    "Pfeiffer Orders",
                    "Busch Orders",
                    "Key Account Orders"
                ],
                "Resource Hours": [
                    round(receiving_hours, 1),
                    round(pfeiffer_hours, 1),
                    round(busch_hours, 1),
                    round(key_account_hours, 1)
                ]
            })

            fig_capacity = px.bar(
                capacity_breakdown,
                x="Workstream",
                y="Resource Hours",
                text="Resource Hours",
                title="Resource Hours Required by Workstream"
            )

            fig_capacity.update_traces(
                texttemplate="%{text:.1f}",
                textposition="outside"
            )

            st.plotly_chart(fig_capacity, use_container_width=True)

            st.dataframe(
                capacity_breakdown,
                use_container_width=True,
                hide_index=True
            )

    # =====================================================
    # TAB 3 — ABOUT RDC
    # =====================================================

    with tab3:

        st.header("About RDC")

        st.subheader("Busch Group Regional Distribution Center")

        st.info("Address: 1910 Campostella Road, Chesapeake, VA 23324")

        st.info("Pickup and Drop-Off Hours: Monday to Friday, 7:00 AM to 3:00 PM")

        st.markdown("---")

        # =====================================================
        # RDC REFERENCE TABLES
        # =====================================================

        st.header("RDC Reference Tables")

        reference_file_path = get_reference_file_path()

        if os.path.exists(reference_file_path):

            ref_bin_master = pd.read_excel(
                reference_file_path,
                sheet_name="Bin_Master"
            )

            ref_product_master = pd.read_excel(
                reference_file_path,
                sheet_name="Product_Master"
            )

            ref_bin_master.columns = ref_bin_master.columns.str.strip()
            ref_product_master.columns = ref_product_master.columns.str.strip()

            required_bin_cols = [
                "Bin Location",
                "Storage Type",
                "Access Classification",
                "Blocked Bin"
            ]

            required_product_cols = [
                "Material",
                "Material Description",
                "Plant-Sp.Matl Status",
                "X-Plant Material Status",
                "Procurement Type",
                "Country of origin"
            ]

            for col in required_bin_cols:
                if col not in ref_bin_master.columns:
                    ref_bin_master[col] = "Blank"

            for col in required_product_cols:
                if col not in ref_product_master.columns:
                    ref_product_master[col] = "Blank"

            tab_bin, tab_product = st.tabs(
                [
                    "Bin Master Table",
                    "Product Master Table"
                ]
            )

            with tab_bin:

                st.subheader("Bin Master Table")

                st.dataframe(
                    ref_bin_master[required_bin_cols],
                    use_container_width=True,
                    height=600,
                    hide_index=True
                )

            with tab_product:

                st.subheader("Product Master Table")

                st.dataframe(
                    ref_product_master[required_product_cols],
                    use_container_width=True,
                    height=600,
                    hide_index=True
                )

        else:

            st.warning(
                f"""
                RDC_Master_Reference.xlsx file not found.

                Expected location:
                {reference_file_path}
                """
            )

        st.markdown("---")

        # =====================================================
        # WAREHOUSE RECOMMENDATIONS
        # =====================================================

        st.header("Warehouse Handling and Process Recommendations")

        st.subheader("Receiving Recommendations")

        st.write(
            """
            - If a product or shipment is physically damaged, immediately inform the supervisor.
            - Verify product number, quantity, labels, serial number barcode if applicable, and visible condition.
            - For Pfeiffer products, make sure there are no supplier labels or supplier cartons unless specifically allowed.
            - Some sealed supplier packaging may not need to be opened. When unsure, check with QA or technical support.
            - If supplier serial number barcodes or product labels are incorrect, inform the supervisor immediately.
            - Proper material handling is critical because cosmetic damage can create customer complaints and nonconformance risk.
            """
        )

        st.subheader("Putaway and Storage Recommendations")

        st.write(
            """
            - Use the correct bin size and storage space for each product.
            - Do not overcrowd products in bins.
            - Avoid placing heavy items in small bins when there is handling or damage risk.
            - Handle products carefully during movement.
            - Use suitable locations based on product size, weight, pick frequency, and access classification.
            """
        )

        st.subheader("Picking Recommendations")

        st.write(
            """
            - At the bin location, verify the product number against the pick list.
            - Confirm the correct quantity before moving the product forward.
            - Be careful with product labels, supplier labels, and cartons during picking.
            - If anything does not match the pick list, escalate before continuing.
            """
        )

        st.subheader("Packing and Shipment Processing Recommendations")

        st.write(
            """
            - Packing should verify product number, quantity, labels, and carton condition again.
            - Heavy items should be packaged carefully to reduce freight damage risk.
            - Make sure all packaging instructions and additional documentation requirements are addressed before final packing.
            - If packaging instructions are missing, use the SharePoint process to request guidance.
            - If there is no response within 24 hours, escalate to the supervisor.
            """
        )

        st.subheader("Outbound QA Recommendations")

        st.write(
            """
            - When performing outbound QA using the app, take clear pictures.
            - Product labels and quantities should be visible as much as possible.
            - Picking, packing, and QA are all responsible for confirming the product is correct before shipment.
            """
        )

        st.subheader("Key Account Shipment Recommendations")

        st.write(
            """
            - For AMAT, LAM, and ASM shipments, maintain at least three trained associates deep where possible.
            - Use the key account Teams channel for fast escalation.
            - If anything unusual occurs in a key account portal, escalate immediately.
            """
        )

        st.markdown("---")

        # =====================================================
        # RDC FAQ
        # =====================================================

        st.header("RDC FAQ")

        with st.expander("What should I do if material is physically damaged during receiving?"):
            st.write("Immediately inform the supervisor before continuing the receiving process.")

        with st.expander("What should I verify during receiving?"):
            st.write("Verify product number, quantity, labels, serial number barcode if applicable, and visible product condition.")

        with st.expander("What should I do if supplier labels or barcodes are incorrect?"):
            st.write("Escalate to the supervisor immediately. Do not continue processing without clarification.")

        with st.expander("What should I check during picking?"):
            st.write("Verify the product number and quantity against the pick list at the bin location before moving the product forward.")

        with st.expander("What should packing verify?"):
            st.write("Packing should verify product number, quantity, labels, carton condition, and any packaging instructions before shipment.")

        with st.expander("What should QA capture in outbound pictures?"):
            st.write("QA should capture clear images where the product label and quantity are visible as much as possible.")

        with st.expander("What should we do if packaging instructions are missing?"):
            st.write("Use the SharePoint process to request instructions. If there is no response within 24 hours, escalate to the supervisor.")

        with st.expander("What is important for AMAT, LAM, and ASM shipments?"):
            st.write("Maintain trained backup coverage and escalate immediately if anything unusual occurs in the key account portal.")

    # =====================================================
    # TAB 4 — HISTORICAL RUNS
    # =====================================================

    with tab4:

        st.header("Historical Runs")

        history_folder = "Historical_Runs"

        if os.path.exists(history_folder):

            folders = [
                f for f in os.listdir(history_folder)
                if os.path.isdir(os.path.join(history_folder, f))
            ]

            if folders:

                selected_folder = st.selectbox(
                    "Select Historical Run",
                    sorted(folders, reverse=True)
                )

                selected_path = os.path.join(history_folder, selected_folder)

                files = os.listdir(selected_path)

                for file in files:

                    file_path = os.path.join(selected_path, file)

                    with open(file_path, "rb") as f:

                        st.download_button(
                            label=f"Download {file}",
                            data=f.read(),
                            file_name=file
                        )

            else:

                st.info("No historical runs saved.")

        else:

            st.info("No historical runs folder found.")