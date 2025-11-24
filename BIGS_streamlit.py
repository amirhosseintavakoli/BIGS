import pandas as pd
import numpy as np
import streamlit as st
import matplotlib.pyplot as plt
from matplotlib.patches import Patch


# Small, focused Streamlit app: load data once, plot on selection changes.

outcome_label2 = {'value': '% Support Value', 'num': '% Enterprise', 'avg_value': '% Average Support'}
color_mapping = {
    'beneficiary': '#1f77b4', 'support_type': '#ff7f0e', 'enterprise_type': '#2ca02c', 'age': '#d62728',
    'support_intensity': '#9467bd', 'province': '#8c564b', 'industry': '#e377c2', 'export': '#7f7f7f',
    'rd': '#bcbd22', 'support_rev': '#17becf', 'emp': '#aec7e8', 'hg': '#ffbb78'
}
labels = {
    'beneficiary': 'Beneficiary Type', 'support_type': 'Support Type', 'enterprise_type': 'Enterprise Type',
    'age': 'Firm Age', 'support_intensity': 'Support', 'province': 'Region', 'industry': 'Industry',
    'export': 'Export Status', 'rd': 'R&D Status', 'support_rev': 'Support as a Proportion of Revenue',
    'emp': 'Firm Size by Employment', 'hg': 'High-Growth (HG) by Revenue'
}
program_select_list = ['CSBFP (ISED)', 'SIF (ISED)', 'IRAP (NRC)','Mitacs (ISED)', 'Alliance Grant (NSERC)', 'SREP (NRCAN)']
outcome_list = ['value', 'num', 'avg_value']
year_list = list(range(2022, 2014, -1))


@st.cache_data
def load_bigs_data(path='BIGS_ProgramStream_Tables_2014_2022_Final.xlsx'):
    """Read all tables into a dict of DataFrames. Cached to run once per path."""
    dfs = {}

    # The original file used many sheets; keep the same names and header offsets.
    # Minimal renames to unify column names.
    df = pd.read_excel('BIGS_ProgramStream_Tables_2014_2022_Final.xlsx',sheet_name='Table 1',header=3,skipfooter=7)
    df = df.rename(columns={'Year of\nSupport': 'year',
                            'Program ID': 'program_id',
                            'Program': 'program',
                            'Province/territory':'Province',
                            'Type of support': 'type',
                            'Value of Support\nLevel': 'value_level',
                            'Type of\nbeneficiary': 'type',
                            'Number of\nultimate\nbeneficiary\nentreprises': 'num',
                            'Total value\nof support to\nultimate\nbeneficiary\nenterprises\n(rounded to\nnearest\nthousands\ndollars)': 'value'})
    dfs['beneficiary'] = df

    df = pd.read_excel('BIGS_ProgramStream_Tables_2014_2022_Final.xlsx',sheet_name='Table 2',header=3,skipfooter=10)
    df = df.rename(columns={'Year of\nSupport': 'year',
                            'Program ID': 'program_id',
                            'Program': 'program',
                            'Province/territory':'Province',
                            'Type of support': 'type',
                            'Value of Support\nLevel': 'value_level',
                            'Type of\nbeneficiary': 'beneficiary',
                            'Number of\nultimate\nbeneficiary\nentreprises': 'num',
                            'Total value of\nsupport to\nultimate\nbeneficiary\nenterprises\n(rounded to\nnearest\nthousands\ndollars)': 'value'})    
    dfs['support_type'] = df

    df = pd.read_excel('BIGS_ProgramStream_Tables_2014_2022_Final.xlsx',sheet_name='Table 3',header=3,skipfooter=7)
    df = df.rename(columns={'Year of\nSupport': 'year',
                            'Program ID': 'program_id',
                            'Program': 'program',
                            'Province/territory':'Province',
                            'Type of support': 'type',
                            'Type of enterprise': 'type',
                            'Value of Support\nLevel': 'value_level',
                            'Type of\nbeneficiary': 'beneficiary',
                            'Number of\nultimate\nbeneficiary\nentreprises': 'num',
                            'Total value of\nsupport to\nultimate\nbeneficiary\nenterprises\n(rounded to\nnearest\nthousands\ndollars)': 'value'})
    dfs['enterprise_type'] = df

    df = pd.read_excel('BIGS_ProgramStream_Tables_2014_2022_Final.xlsx',sheet_name='Table 4',header=3,skipfooter=7)
    df = df.rename(columns={'Year of\nSupport': 'year',
                            'Program ID': 'program_id',
                            'Program': 'program',
                            'Province/territory':'Province',
                            'Type of support': 'type',
                            'Type of enterprise': 'type',
                            'Value of Support\nLevel': 'value_level',
                            'Type of\nbeneficiary': 'beneficiary',
                            'Firm age': 'type',
                            'Number of\nultimate\nbeneficiary\nentreprises': 'num',
                            'Total value\nof support to\nultimate\nbeneficiary\nenterprises\n(rounded to\nnearest\nthousands\ndollars)': 'value'})
    dfs['age'] = df

    df = pd.read_excel('BIGS_ProgramStream_Tables_2014_2022_Final.xlsx',sheet_name='Table 5',header=3,skipfooter=7)
    df = df.rename(columns={'Year of\nSupport': 'year',
                            'Program ID': 'program_id',
                            'Program': 'program',
                            'Province/territory':'Province',
                            'Type of support': 'type',
                            'Type of enterprise': 'type',
                            'Value of support': 'type',
                            'Value of Support\nLevel': 'value_level',
                            'Type of\nbeneficiary': 'beneficiary',
                            'Firm age': 'age',
                            'Number of\nultimate\nbeneficiary\nentreprises': 'num',
                            'Total value of\nsupport to\nultimate\nbeneficiary\nenterprises\n(rounded to\nnearest\nthousands\ndollars)': 'value',
                            'Total value\nof support to\nultimate\nbeneficiary\nenterprises\n(rounded to\nnearest\nthousands\ndollars)': 'value'})
    dfs['support_intensity'] = df

    df = pd.read_excel('BIGS_ProgramStream_Tables_2014_2022_Final.xlsx',sheet_name='Table 6',header=3,skipfooter=7)
    df = df.rename(columns={'Year of\nSupport': 'year',
                            'Program ID': 'program_id',
                            'Program': 'program',
                            'Province/territory':'Province',
                            'Type of support': 'type',
                            'Type of enterprise': 'type',
                            'Value of support': 'type',
                            'Region': 'type',
                            'Value of Support\nLevel': 'value_level',
                            'Type of\nbeneficiary': 'beneficiary',
                            'Firm age': 'age',
                            'Number of\nultimate\nbeneficiary\nentreprises': 'num',
                            'Total value of\nsupport to\nultimate\nbeneficiary\nenterprises\n(rounded to\nnearest\nthousands\ndollars)': 'value',
                            'Total value\nof support to\nultimate\nbeneficiary\nenterprises\n(rounded to\nnearest\nthousands\ndollars)': 'value'})

    dfs['province'] = df

    df = pd.read_excel('BIGS_ProgramStream_Tables_2014_2022_Final.xlsx',sheet_name='Table 7',header=3,skipfooter=8)
    df = df.rename(columns={'Year of\nSupport': 'year',
                            'Program ID': 'program_id',
                            'Program': 'program',
                            'Province/territory':'Province',
                            'Type of support': 'type',
                            'Type of enterprise': 'type',
                            'Value of support': 'type',
                            'Region': 'type',
                            'Industry': 'type',
                            'Value of Support\nLevel': 'value_level',
                            'Type of\nbeneficiary': 'beneficiary',
                            'Firm age': 'age',
                            'Number of\nultimate\nbeneficiary\nentreprises': 'num',
                            'Total value of\nsupport to\nultimate\nbeneficiary\nenterprises\n(rounded to\nnearest\nthousands\ndollars)': 'value',
                            'Total value\nof support to\nultimate\nbeneficiary\nenterprises\n(rounded to\nnearest\nthousands\ndollars)': 'value'})
    dfs['industry'] = df

    df = pd.read_excel('BIGS_ProgramStream_Tables_2014_2022_Final.xlsx',sheet_name='Table 8',header=3,skipfooter=8)
    df = df.rename(columns={'Year of\nSupport': 'year',
                            'Program ID': 'program_id',
                            'Program': 'program',
                            'Province/territory':'Province',
                            'Type of support': 'type',
                            'Type of enterprise': 'type',
                            'Value of support': 'type',
                            'Region': 'type',
                            'Industry': 'type',
                            'R&D\nperformer\nstatus': 'type',
                            'Value of Support\nLevel': 'value_level',
                            'Type of\nbeneficiary': 'beneficiary',
                            'Firm age': 'age',
                            'Number of\nultimate\nbeneficiary\nentreprises': 'num',
                            'Total value of\nsupport to\nultimate\nbeneficiary\nenterprises\n(rounded to\nnearest\nthousands\ndollars)': 'value',
                            'Total value\nof support to\nultimate\nbeneficiary\nenterprises\n(rounded to\nnearest\nthousands\ndollars)': 'value'})
    dfs['rd'] = df

    df = pd.read_excel('BIGS_ProgramStream_Tables_2014_2022_Final.xlsx',sheet_name='Table 9',header=3,skipfooter=7)
    df = df.rename(columns={'Year of\nSupport': 'year',
                            'Program ID': 'program_id',
                            'Program': 'program',
                            'Province/territory':'Province',
                            'Type of support': 'type',
                            'Type of enterprise': 'type',
                            'Value of support': 'type',
                            'Region': 'type',
                            'Industry': 'type',
                            'R&D\nperformer\nstatus': 'type',
                            'Exporter\nstatus': 'type',
                            'Value of Support\nLevel': 'value_level',
                            'Type of\nbeneficiary': 'beneficiary',
                            'Firm age': 'age',
                            'Number of\nultimate\nbeneficiary\nentreprises': 'num',
                            'Total value of\nsupport to\nultimate\nbeneficiary\nenterprises\n(rounded to\nnearest\nthousands\ndollars)': 'value',
                            'Total value\nof support to\nultimate\nbeneficiary\nenterprises\n(rounded to\nnearest\nthousands\ndollars)': 'value'})
    dfs['export'] = df

    df = pd.read_excel('BIGS_ProgramStream_Tables_2014_2022_Final.xlsx',sheet_name='Table 10',header=3,skipfooter=7)
    df = df.rename(columns={'Year of\nSupport': 'year',
                            'Program ID': 'program_id',
                            'Program': 'program',
                            'Province/territory':'Province',
                            'Type of support': 'type',
                            'Type of enterprise': 'type',
                            'Value of support': 'type',
                            'Region': 'type',
                            'Industry': 'type',
                            'R&D\nperformer\nstatus': 'type',
                            'Exporter\nstatus': 'type',
                            'Support as a\nproportion of revenue': 'type',
                            'Value of Support\nLevel': 'value_level',
                            'Type of\nbeneficiary': 'beneficiary',
                            'Firm age': 'age',
                            'Number of\nultimate\nbeneficiary\nentreprises': 'num',
                            'Total value of\nsupport to\nultimate\nbeneficiary\nenterprises\n(rounded to\nnearest\nthousands\ndollars)': 'value',
                            'Total value\nof support to\nultimate\nbeneficiary\nenterprises\n(rounded to\nnearest\nthousands\ndollars)': 'value'})
    dfs['support_rev'] = df

    df = pd.read_excel('BIGS_ProgramStream_Tables_2014_2022_Final.xlsx',sheet_name='Table 11',header=3,skipfooter=8)
    df = df.rename(columns={'Year of\nSupport': 'year',
                            'Program ID': 'program_id',
                            'Program': 'program',
                            'Province/territory':'Province',
                            'Type of support': 'type',
                            'Type of enterprise': 'type',
                            'Value of support': 'type',
                            'Region': 'type',
                            'Industry': 'type',
                            'R&D\nperformer\nstatus': 'type',
                            'Exporter\nstatus': 'type',
                            'Support as a\nproportion of revenue': 'type',
                            'Employment\nsize': 'type',
                            'Value of Support\nLevel': 'value_level',
                            'Type of\nbeneficiary': 'beneficiary',
                            'Firm age': 'age',
                            'Number of\nultimate\nbeneficiary\nentreprises': 'num',
                            'Total value of\nsupport to\nultimate\nbeneficiary\nenterprises\n(rounded to\nnearest\nthousands\ndollars)': 'value',
                            'Total value\nof support to\nultimate\nbeneficiary\nenterprises\n(rounded to\nnearest\nthousands\ndollars)': 'value'})
    dfs['emp'] = df

    df = pd.read_excel('BIGS_ProgramStream_Tables_2014_2022_Final.xlsx',sheet_name='Table 12',header=3,skipfooter=8)
    df = df.rename(columns={'Year of\nSupport': 'year',
                            'Program ID': 'program_id',
                            'Program': 'program',
                            'Province/territory':'Province',
                            'Type of support': 'type',
                            'Type of enterprise': 'type',
                            'Value of support': 'type',
                            'Region': 'type',
                            'Industry': 'type',
                            'R&D\nperformer\nstatus': 'type',
                            'Exporter\nstatus': 'type',
                            'Support as a\nproportion of revenue': 'type',
                            'Employment\nsize': 'type',
                            'High-growth-by-revenue\nstatus': 'type',
                            'Value of Support\nLevel': 'value_level',
                            'Type of\nbeneficiary': 'beneficiary',
                            'Firm age': 'age',
                            'Number of\nultimate\nbeneficiary\nentreprises': 'num',
                            'Total value of\nsupport to\nultimate\nbeneficiary\nenterprises\n(rounded to\nnearest\nthousands\ndollars)': 'value',
                            'Total value\nof support to\nultimate\nbeneficiary\nenterprises\n(rounded to\nnearest\nthousands\ndollars)': 'value'})
    dfs['hg'] = df
    return dfs


def clean_bigs_data(df):
    # Normalize numeric columns and create avg value
    df['value'] = df['value'].replace(['X', 'x', '...'], np.nan)
    df['value'] = df['value'].astype(str).str.replace(',', '').astype(float) / 1e6
    df['num'] = df['num'].replace(['X', 'x', '...'], np.nan)
    df['num'] = df['num'].astype(str).str.replace(',', '').astype(float)
    df['avg_value'] = df['value'] / df['num'] * 1e3

    # Clean strings
    for col in ['program', 'value_level', 'type']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace('\n', ' ')

    # Create totals dataframe
    df2 = df[df.get('program_id', pd.Series()).astype(str).str.contains('TOTAL', na=False)].copy()
    if 'value_level' in df2.columns:
        df2 = df2[df2['value_level'].str.contains('Total', na=False)]
        df2 = df2.drop(columns=['value_level'], errors='ignore')

    # Filter program-level rows and tidy program ids
    mask = df.get('program_id', pd.Series()).astype(str).str.contains('Total|_00|IRAP', na=False)
    df = df[mask].copy()
    df = df[~((df.get('program', pd.Series()).str.contains('NRC_CNRC', na=False)) & (~df.get('program_id', pd.Series()).str.contains('IRAP', na=False)))]
    df['program_id'] = df.get('program_id', pd.Series()).astype(str).str.replace('_00', '', regex=False).str.replace('(Total)', '', regex=False).str.strip()
    df['program'] = df.get('program', pd.Series()).astype(str).str.replace('(Total)', '', regex=False).str.strip()

    if 'value_level' in df.columns:
        df = df[df['value_level'].str.contains('Total', na=False)].copy()
        df = df.drop(columns=['value_level'], errors='ignore')

    df = df.drop_duplicates(subset=['year', 'program_id', 'type'], keep='first')

    # Program select labels
    df['program_select'] = df.get('program', pd.Series())
    df.loc[df['program'].str.contains('Canada Small Business Financing Program', na=False), 'program_select'] = 'CSBFP (ISED)'
    df.loc[df['program'].str.contains('Strategic Innovation Fund', na=False), 'program_select'] = 'SIF (ISED)'
    df.loc[df['program_id'].str.match('IRAP', na=False), 'program_select'] = 'IRAP (NRC)'
    df.loc[df['program'].str.contains('Mitacs', na=False), 'program_select'] = 'Mitacs (ISED)'
    df.loc[df['program'].str.contains('Alliance grant', na=False ), 'program_select'] = 'Alliance Grant (NSERC)'
    df.loc[df['program'].str.contains('Smart Renewables', na=False), 'program_select'] = 'SREP (NRCAN)'    

    return df, df2


def bargraph_labelling(ax, df, df_total, outcome, total, total_value, key):
    if df.empty:
        return
    df = df.copy()
    df = df.dropna(subset=[outcome])
    df['percentage'] = df[outcome] / total * 100
    grouped = df_total.groupby('type').agg(total_value=(outcome, 'sum')) if not df_total.empty else pd.DataFrame()
    if not grouped.empty and total_value != 0:
        grouped['percentage'] = grouped['total_value'] / total_value * 100

    bars = ax.barh(df['type'], df['percentage'], color=color_mapping.get(key, 'C0'), alpha=0.7, label=labels.get(key, key))
    for bar, pct in zip(bars, df['percentage']):
        ax.text(bar.get_width() + 1, bar.get_y() + bar.get_height() / 2, f'{pct:.1f}%', va='center', fontsize=8)
    
    # add the category's name to the figure
    gap_df = pd.DataFrame([{'type': labels[key],
                            'percentage': 0}])

    grouped = grouped[grouped.index.isin(df['type'])]
    ax.barh(grouped.index, grouped['percentage'], facecolor='none', edgecolor='dimgray', linestyle='--', linewidth=1.5)
    ax.barh(gap_df['type'], gap_df['percentage'])


def plot_bigs_program_stream(ax, dataframes, dfs_total, program, outcome, year):
    # total value of support for each program-year
    df_main = dataframes.get('support_type', pd.DataFrame())
    df_main = df_main[(df_main.get('program_select') == program) & (df_main.get('year') == year) & (df_main.get('type') == 'Total')]
    total = df_main[outcome].sum() if not df_main.empty else 0

    # total value of support in each year
    df_total_main = dfs_total.get('support_type', pd.DataFrame())
    df_total_main = df_total_main[(df_total_main.get('year') == year) & (df_total_main.get('type') == 'Total')]
    total_value = df_total_main[outcome].values[0] if not df_total_main.empty else 0

    category_list = ['support_intensity', 'province', 'industry', 'emp', 'age', 'rd', 'export', 'support_rev', 'hg']
    for key in category_list[::-1]:
        df_cat = dataframes.get(key, pd.DataFrame())
        df_cat = df_cat[(df_cat.get('program_select') == program) & (df_cat.get('year') == year)]
        df_total_cat = dfs_total.get(key, pd.DataFrame())
        df_total_cat = df_total_cat[df_total_cat.get('year') == year] if not df_total_cat.empty else pd.DataFrame()

        # normalize some category labels for readability (examples)
        if key == 'support_intensity' and not df_cat.empty:
            df = df_cat[~df_cat.type.isin(['Total'])].copy()
            df.loc[df['type'].str.contains('less than', na=False), 'type'] = '< 100k'
            df.loc[df['type'].str.contains('between', na=False), 'type'] = '100k - 1mil'
            df.loc[df['type'].str.contains('more than', na=False), 'type'] = '> 1mil'
            df_total = df_total_cat[~df_total_cat.type.isin(['Total'])].copy()
            df_total.loc[df_total['type'].str.contains('less than', na=False), 'type'] = '< 100k'
            df_total.loc[df_total['type'].str.contains('between', na=False), 'type'] = '100k - 1mil'
            df_total.loc[df_total['type'].str.contains('more than', na=False), 'type'] = '> 1mil'
            bargraph_labelling(ax, df, df_total, outcome, total, total_value, key)
        elif key == 'emp' and not df_cat.empty:
            df = df_cat[df_cat.type.isin(['Large enterprises',
                                'Medium Enterprises',
                                'Small enterprises'])].copy()
            df.loc[df['type'].str.contains('Small enterprises'), 'type'] = '< 100'
            df.loc[df['type'].str.contains('Medium Enterprises'), 'type'] = '100-499'
            df.loc[df['type'].str.contains('Large enterprises'), 'type'] = '>= 500'
            df_total = df_total_cat[df_total_cat.type.isin(['Large enterprises',
                                'Medium Enterprises',
                                'Small enterprises'])].copy()
            df_total.loc[df_total['type'].str.contains('Small enterprises'), 'type'] = '< 100'
            df_total.loc[df_total['type'].str.contains('Medium Enterprises'), 'type'] = '100-499'
            df_total.loc[df_total['type'].str.contains('Large enterprises'), 'type'] = '>= 500'
            bargraph_labelling(ax, df, df_total, outcome, total, total_value, key)
        elif key == 'age' and not df_cat.empty:
            df = df_cat[df_cat.type.isin(['More than 20 years old', '11-20 years old', '6-10 years old', '2-5 years old', '1 year old or less'])].copy()
            df_total = df_total_cat[df_total_cat.type.isin(['More than 20 years old', '11-20 years old', '6-10 years old', '2-5 years old', '1 year old or less'])].copy() if not df_total_cat.empty else pd.DataFrame()
            bargraph_labelling(ax, df, df_total, outcome, total, total_value, key)
        elif key == 'province' and not df_cat.empty:
            df = df_cat[df_cat.type.isin(['British Columbia', 'Prairies', 'Ontario', 'Quebec', 'Atlantic', 'Territories'])].copy()
            df_total = df_total_cat[df_total_cat.type.isin(['British Columbia', 'Prairies', 'Ontario', 'Quebec', 'Atlantic', 'Territories'])].copy() if not df_total_cat.empty else pd.DataFrame()
            bargraph_labelling(ax, df, df_total, outcome, total, total_value, key)
        elif key == 'industry' and not df_cat.empty:
            df = df_cat[df_cat.type.isin(['All other industries',
                              '54 - Professional, Scientific and Technical Services',
                              '44-45 - Wholesale and Retail Trade',
                              '31-33 - Manufacturing'])].copy()
            df.loc[df['type'].str.contains('All other industries'), 'type'] = 'Other'
            df.loc[df['type'].str.contains('Professional'), 'type'] = 'Prof/Sci/Tech'
            df.loc[df['type'].str.contains('Wholesale and Retail'), 'type'] = 'Wholesale/Retail'
            df.loc[df['type'].str.contains('Manufacturing'), 'type'] = 'Manufacturing'
            df_total = df_total_cat[df_total_cat.type.isin(['All other industries',
                              '54 - Professional, Scientific and Technical Services',
                              '44-45 - Wholesale and Retail Trade',
                              '31-33 - Manufacturing'])].copy() if not df_total_cat.empty else pd.DataFrame()
            df_total.loc[df_total['type'].str.contains('All other industries'), 'type'] = 'Other'
            df_total.loc[df_total['type'].str.contains('Professional'), 'type'] = 'Prof/Sci/Tech'
            df_total.loc[df_total['type'].str.contains('Wholesale and Retail'), 'type'] = 'Wholesale/Retail'
            df_total.loc[df_total['type'].str.contains('Manufacturing'), 'type'] = 'Manufacturing'
            bargraph_labelling(ax, df, df_total, outcome, total, total_value, key)
        elif key == 'export' and not df_cat.empty:
            df  = df_cat[df_cat.type.isin(['Exporter'])].copy()
            df_total = df_total_cat[df_total_cat.type.isin(['Exporter'])].copy()
            bargraph_labelling(ax, df, df_total, outcome, total, total_value, key)
        elif key == 'rd' and not df_cat.empty:
            df  = df_cat[df_cat.type.isin(['R&D performer'])].copy()
            df_total = df_total_cat[df_total_cat.type.isin(['R&D performer'])].copy()
            bargraph_labelling(ax, df, df_total, outcome, total, total_value, key)
        elif key == 'support_rev' and not df_cat.empty:
            df = df_cat[df_cat.type.isin(['Value of support received is more than 50% of revenue',
                                'Value of support received is more than 25% of revenue and 50% of revenue or less',
                                'Value of support received is more than 10% of revenue and 25% of revenue or less',
                                'Value of support received is 10% of revenue or less'])].copy()
            df.loc[df['type'].str.contains('10% of revenue or less'), 'type'] = '< 10%'
            df.loc[df['type'].str.contains('25% of revenue or less'), 'type'] = '10%-25%'
            df.loc[df['type'].str.contains('50% of revenue or less'), 'type'] = '25%-50%'
            df.loc[df['type'].str.contains('more than 50% of revenue'), 'type'] = '>50%'
            df_total = df_total_cat[df_total_cat.type.isin(['Value of support received is more than 50% of revenue',
                                'Value of support received is more than 25% of revenue and 50% of revenue or less',
                                'Value of support received is more than 10% of revenue and 25% of revenue or less',
                                'Value of support received is 10% of revenue or less'])].copy()
            df_total.loc[df_total['type'].str.contains('10% of revenue or less'), 'type'] = '< 10%'
            df_total.loc[df_total['type'].str.contains('25% of revenue or less'), 'type'] = '10%-25%'
            df_total.loc[df_total['type'].str.contains('50% of revenue or less'), 'type'] = '25%-50%'
            df_total.loc[df_total['type'].str.contains('more than 50% of revenue'), 'type'] = '>50%'
            bargraph_labelling(ax, df, df_total, outcome, total, total_value, key)
        elif key == 'hg' and not df_cat.empty:
            df2 = df_cat[df_cat.type.isin(['Total'])].copy()
            df = df_cat[df_cat.type.isin(['High-growth-by-revenue'])].copy().sort_values(by=outcome)
            df.loc[df['type'].str.contains('High-growth-by-revenue'), 'type'] = 'HG'

            df_total = df_total_cat[df_total_cat.type.isin(['High-growth-by-revenue'])].copy().sort_values(by=outcome)
            df_total.loc[df_total['type'].str.contains('High-growth-by-revenue'), 'type'] = 'HG'

            bargraph_labelling(ax, df, df_total, outcome, total, total_value, key)
        else:
            continue
    
    # Create legend handles using consistent colors and labels
    legend_handles = [
        Patch(facecolor=color_mapping[key], label=labels[key])
        for key in category_list if key in labels and key in color_mapping
    ]
    extra_handle = Patch(facecolor='none', edgecolor='dimgray', linestyle='--', linewidth=1.5, label='BIGS')
    legend_handles.append(extra_handle)
    plt.legend(loc='lower right', handles=legend_handles, fontsize = 5)
    # function does not create the figure; plotting caller should create and display it
    return


st.set_page_config(page_title="BIGS Interactive Stats Viewer", layout="wide")
st.title("Interactive Stats Viewer")


@st.cache_data
def load_and_clean_bigs_data(path='BIGS_ProgramStream_Tables_2014_2022_Final.xlsx'):
    raw = load_bigs_data(path)
    print("Data loaded from Excel.")
    dataframes_clean = {}
    dataframes_total = {}
    for key, df in raw.items():
        dataframes_clean[key], dataframes_total[key] = clean_bigs_data(df.copy())
    print("Data cleaned and processed.")
    return dataframes_clean, dataframes_total


if __name__ == '__main__':
    dataframes_clean, dataframes_total = load_and_clean_bigs_data('BIGS_ProgramStream_Tables_2014_2022_Final.xlsx')

    st.sidebar.markdown('**Controls**')
    program = st.sidebar.selectbox('Select Program:', program_select_list)
    year = st.sidebar.selectbox('Select Year:', year_list)
    outcome = 'value' # default outcome

    st.subheader(f'Distribution of {outcome_label2[outcome]} for {program} in {year}')
    
    # df_main = dataframes_clean.get('support_type', pd.DataFrame())
    # st.write(df_main[(df_main.get('program_select') == program) & (df_main.get('year') == year) & (df_main.get('type') == 'Total')])

    fig, ax = plt.subplots(figsize=(10, 5))
    plot_bigs_program_stream(ax, dataframes_clean, dataframes_total, program, outcome, year)

    # formatting y-axis
    yticks      = ax.get_yticks()
    yticklabels = ax.get_yticklabels()

    # Set bold font for specific categories
    for label in yticklabels:
        if label.get_text() in labels.values():
            label.set_fontweight('bold')

    ax.set_yticks(yticks)
    ax.set_yticklabels(yticklabels)
    plt.yticks(fontsize = 8)
    plt.xticks(fontsize = 8)

    # Remove top and right spines
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    plt.xlabel(outcome_label2[outcome])
    plt.xlim(0, 100)
    plt.grid(False)
    plt.tight_layout()
    st.write(fig)
