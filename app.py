import ast
import io 
import re 
from datetime import datetime 
from uuid import uuid4 
 
import numpy as np 
import pandas as pd 
import streamlit as st 
import matplotlib.pyplot as plt 
 
from openpyxl import load_workbook 
from openpyxl.drawing.image import Image as XLImage 
 
 
# ============================================================ 
# OPTIONAL DOTENV 
# ============================================================ 
 
try: 
    from dotenv import load_dotenv 
    load_dotenv() 
except Exception: 
    pass 
 
 
# ============================================================ 
# PAGE CONFIG 
# ============================================================ 
 
st.set_page_config( 
    page_title="AI Excel Analyst", 
    page_icon="📊", 
    layout="wide" 
) 
 
 
# ============================================================ 
# SESSION STATE 
# ============================================================ 
 
DEFAULT_STATE = { 
    "stage": "upload", 
 
    "raw_df": None, 
    "cleaned_df": None, 
    "transformed_df": None, 
 
    "uploaded_filename": "", 
 
    "inspection": None, 
 
    "cleaning_log": [], 
    "transformation_log": [], 
    "analysis_log": [], 
 
    "validation_df": None, 
 
    "pivot_df": None, 
 
    "analysis_results": [], 
    "report_items": [], 
 
    "chat_messages": [], 
 
    "current_result": None, 
 
    "last_chart": None, 
    "last_chart_name": "", 
    "last_chart_data": None, 
 
    "report_generated": False, 
 
    # Columns created by AI 
    "calculated_columns": [], 
 
    # Protected/source columns 
    "original_columns": [], 
} 
 
 
for key, value in DEFAULT_STATE.items(): 
 
    if key not in st.session_state: 
        st.session_state[key] = value 
 
 
# ============================================================ 
# GENERAL HELPERS 
# ============================================================ 
 
def add_message(role, content): 
 
    st.session_state.chat_messages.append({ 
        "role": role, 
        "content": content 
    }) 
 
 
def clean_text(value): 
 
    if pd.isna(value): 
        return value 
 
    value = str(value) 
 
    value = re.sub( 
        r"\s+", 
        " ", 
        value 
    ) 
 
    return value.strip() 
 
 
def normalize_column_name(name): 
 
    name = str(name) 
 
    name = name.strip() 
 
    name = re.sub( 
        r"\s+", 
        "_", 
        name 
    ) 
 
    name = re.sub( 
        r"[^\w]+", 
        "_", 
        name 
    ) 
 
    name = re.sub( 
        r"_+", 
        "_", 
        name 
    ) 
 
    name = name.strip("_") 
 
    return name 
 
 
def get_working_df(): 
 
    if st.session_state.transformed_df is not None: 
        return st.session_state.transformed_df 
 
    if st.session_state.cleaned_df is not None: 
        return st.session_state.cleaned_df 
 
    return st.session_state.raw_df 
 
 
def find_column(df, requested_name): 
 
    if df is None: 
        return None 
 
    requested = str( 
        requested_name 
    ).strip().lower() 
 
    for col in df.columns: 
 
        if ( 
            str(col).strip().lower() 
            == 
            requested 
        ): 
            return col 
 
    return None 
 
 
def find_column_in_list( 
    columns, 
    requested_name 
): 
 
    requested = str( 
        requested_name 
    ).strip().lower() 
 
    for col in columns: 
 
        if ( 
            str(col).strip().lower() 
            == 
            requested 
        ): 
            return col 
 
    return None 
 
 
def find_numeric_columns(df): 
 
    if df is None: 
        return [] 
 
    return df.select_dtypes( 
        include=np.number 
    ).columns.tolist() 
 
 
def find_text_columns(df): 
 
    if df is None: 
        return [] 
 
    return df.select_dtypes( 
        include=[ 
            "object", 
            "string" 
        ] 
    ).columns.tolist() 
 
 
def find_date_columns(df): 
 
    if df is None: 
        return [] 
 
    result = [] 
 
    for col in df.columns: 
 
        if pd.api.types.is_datetime64_any_dtype( 
            df[col] 
        ): 
 
            result.append(col) 
 
        elif any( 
            word in str(col).lower() 
            for word in [ 
                "date", 
                "time", 
                "timestamp" 
            ] 
        ): 
 
            result.append(col) 
 
    return result 
 
 
# ============================================================ 
# CALCULATED COLUMN HELPERS 
# ============================================================ 
 
def is_calculated_column( 
    column_name 
): 
    """ 
    Checks whether the column was created 
    by the AI during this session. 
    """ 
 
    return any( 
        str(column_name).strip().lower() 
        == 
        str(calculated).strip().lower() 
        for calculated 
        in st.session_state.calculated_columns 
    ) 
 
 
def get_calculated_columns( 
    df=None 
): 
    """ 
    Returns calculated columns that currently 
    exist in the dataframe. 
    """ 
 
    if df is None: 
        df = get_working_df() 
 
    if df is None: 
        return [] 
 
    existing = [] 
 
    for calculated in ( 
        st.session_state.calculated_columns 
    ): 
 
        matching_column = find_column( 
            df, 
            calculated 
        ) 
 
        if matching_column is not None: 
 
            existing.append( 
                matching_column 
            ) 
 
    return existing 
 
 
def find_registered_calculated_column( 
    requested_column 
): 
    """ 
    Finds an AI-created column using 
    case-insensitive matching. 
    """ 
 
    requested_column = str( 
        requested_column 
    ).strip() 
 
    for calculated in ( 
        st.session_state.calculated_columns 
    ): 
 
        if ( 
            str(calculated).strip().lower() 
            == 
            requested_column.lower() 
        ): 
 
            return calculated 
 
    return None 
 
 
def extract_create_calculated_column_request( 
    request 
): 
    """ 
    Converts natural-language creation commands 
    into a simple formula. 
 
    Examples: 
 
    Create a calculated column Revenue = Quantity * Price 
 
    Create calculated column Profit = Revenue - Cost 
 
    Create column Discount = Price * 0.90 
 
    Add calculated column Total = Quantity * Price 
 
    Revenue = Quantity * Price 
    """ 
 
    text = str( 
        request 
    ).strip() 
 
    if not text: 
        return None 
 
    # -------------------------------------------------------- 
    # REMOVE NATURAL-LANGUAGE PREFIX 
    # -------------------------------------------------------- 
 
    patterns = [ 
 
        r"^\s*create\s+(?:a\s+)?calculated\s+column\s+", 
 
        r"^\s*create\s+calculated\s+column\s+", 
 
        r"^\s*create\s+(?:a\s+)?column\s+", 
 
        r"^\s*add\s+(?:a\s+)?calculated\s+column\s+", 
 
        r"^\s*add\s+calculated\s+column\s+", 
 
        r"^\s*add\s+(?:a\s+)?column\s+", 
 
        r"^\s*make\s+(?:a\s+)?calculated\s+column\s+", 
 
        r"^\s*make\s+calculated\s+column\s+", 
 
        r"^\s*calculate\s+column\s+", 
    ] 
 
    cleaned_text = text 
 
    for pattern in patterns: 
 
        new_text = re.sub( 
            pattern, 
            "", 
            cleaned_text, 
            flags=re.IGNORECASE 
        ) 
 
        if new_text != cleaned_text: 
 
            cleaned_text = new_text 
 
            break 
 
    # -------------------------------------------------------- 
    # FORMULA MUST CONTAIN = 
    # -------------------------------------------------------- 
 
    if "=" not in cleaned_text: 
        return None 
 
    return cleaned_text.strip() 
 
 
def extract_delete_column_request( 
    request 
): 
    """ 
    Extracts the column name from delete commands. 
 
    Supported: 
 
    delete revenue 
 
    delete Revenue 
 
    delete calculated column Revenue 
 
    delete the calculated column Revenue 
 
    remove Revenue 
 
    remove calculated column Profit 
 
    delete column Total 
    """ 
 
    text = str( 
        request 
    ).strip() 
 
    if not text: 
        return None 
 
    patterns = [ 
 
        # delete calculated column Revenue 
        r"^\s*(?:delete|remove)\s+" 
        r"(?:the\s+)?" 
        r"(?:calculated\s+|derived\s+)?" 
        r"column\s+(.+?)\s*$", 
 
        # delete calculated Revenue 
        r"^\s*(?:delete|remove)\s+" 
        r"(?:the\s+)?" 
        r"(?:calculated\s+|derived\s+)" 
        r"(.+?)\s*$", 
 
        # delete Revenue 
        r"^\s*(?:delete|remove)\s+" 
        r"(.+?)\s*$", 
    ] 
 
    for pattern in patterns: 
 
        match = re.match( 
            pattern, 
            text, 
            flags=re.IGNORECASE 
        ) 
 
        if match: 
 
            column = match.group( 
                1 
            ).strip() 
 
            # Remove quotes 
            column = column.strip( 
                "`\"'" 
            ) 
 
            # Remove trailing "column" 
            column = re.sub( 
                r"\s+column$", 
                "", 
                column, 
                flags=re.IGNORECASE 
            ).strip() 
 
            if column: 
                return column 
 
    return None 
 
 
def is_delete_column_request( 
    request 
): 
    """ 
    Detects: 
 
    delete revenue 
    remove revenue 
    delete calculated column revenue 
    """ 
 
    text = str( 
        request 
    ).strip() 
 
    return bool( 
        re.match( 
            r"^\s*(delete|remove)\b", 
            text, 
            flags=re.IGNORECASE 
        ) 
    ) 
 
 
def remove_calculated_column( 
    df, 
    requested_column 
): 
    """ 
    Deletes a column ONLY when it was created 
    by the AI. 
 
    Original Excel columns are protected. 
    """ 
 
    if df is None: 
 
        return None, ( 
            "No dataset is available." 
        ) 
 
    requested_column = str( 
        requested_column 
    ).strip() 
 
    if not requested_column: 
 
        return None, ( 
            "Please specify the calculated " 
            "column you want to delete." 
        ) 
 
    # -------------------------------------------------------- 
    # CHECK AI REGISTRY 
    # -------------------------------------------------------- 
 
    registered_column = ( 
        find_registered_calculated_column( 
            requested_column 
        ) 
    ) 
 
    # -------------------------------------------------------- 
    # NOT AN AI COLUMN 
    # -------------------------------------------------------- 
 
    if registered_column is None: 
 
        existing_column = find_column( 
            df, 
            requested_column 
        ) 
 
        if existing_column is not None: 
 
            return None, ( 
                f"'{existing_column}' is an original " 
                f"Excel/source column. Original columns " 
                f"cannot be deleted." 
            ) 
 
        return None, ( 
            f"Calculated column " 
            f"'{requested_column}' was not found." 
        ) 
 
    # -------------------------------------------------------- 
    # FIND ACTUAL DATAFRAME COLUMN 
    # -------------------------------------------------------- 
 
    actual_column = find_column( 
        df, 
        registered_column 
    ) 
 
    if actual_column is None: 
 
        # Remove stale registry item 
        st.session_state.calculated_columns = [ 
            col 
            for col 
            in st.session_state.calculated_columns 
            if ( 
                str(col).strip().lower() 
                != 
                str(registered_column).strip().lower() 
            ) 
        ] 
 
        return None, ( 
            f"Calculated column " 
            f"'{registered_column}' no longer exists." 
        ) 
 
    # -------------------------------------------------------- 
    # DELETE 
    # -------------------------------------------------------- 
 
    result = df.drop( 
        columns=[ 
            actual_column 
        ] 
    ).copy() 
 
    # -------------------------------------------------------- 
    # REMOVE FROM REGISTRY 
    # -------------------------------------------------------- 
 
    st.session_state.calculated_columns = [ 
        col 
        for col 
        in st.session_state.calculated_columns 
        if ( 
            str(col).strip().lower() 
            != 
            str(actual_column).strip().lower() 
        ) 
    ] 
 
    return result, ( 
        f"Deleted calculated column " 
        f"'{actual_column}'." 
    ) 
 
 
# ============================================================ 
# INSPECTION 
# ============================================================ 
 
def inspect_dataframe(df): 
 
    if df is None: 
        return {} 
 
    return { 
        "rows": len(df), 
 
        "columns": len( 
            df.columns 
        ), 
 
        "duplicate_rows": int( 
            df.duplicated().sum() 
        ), 
 
        "missing_cells": int( 
            df.isna() 
            .sum() 
            .sum() 
        ), 
 
        "blank_rows": int( 
            df.isna() 
            .all(axis=1) 
            .sum() 
        ), 
 
        "columns_list": list( 
            df.columns 
        ), 
 
        "dtypes": { 
            str(col): str(dtype) 
            for col, dtype 
            in df.dtypes.items() 
        } 
    } 
 
 
# ============================================================ 
# AUTOMATIC CLEANING 
# ============================================================ 
 
def automatic_cleaning(df): 
 
    result = df.copy() 
 
    log = [] 
 
    original_rows = len( 
        result 
    ) 
 
    # -------------------------------------------------------- 
    # 1. CLEAN COLUMN NAMES 
    # -------------------------------------------------------- 
 
    new_columns = [] 
 
    for col in result.columns: 
 
        new_col = normalize_column_name( 
            col 
        ) 
 
        new_columns.append( 
            new_col 
        ) 
 
    # Make duplicate column names unique 
    seen = {} 
 
    unique_columns = [] 
 
    for col in new_columns: 
 
        if col not in seen: 
 
            seen[col] = 0 
 
            unique_columns.append( 
                col 
            ) 
 
        else: 
 
            seen[col] += 1 
 
            unique_columns.append( 
                f"{col}_{seen[col]}" 
            ) 
 
    if list(result.columns) != ( 
        unique_columns 
    ): 
 
        result.columns = ( 
            unique_columns 
        ) 
 
        log.append( 
            "Column names were cleaned " 
            "and standardized." 
        ) 
 
    # -------------------------------------------------------- 
    # 2. REMOVE BLANK ROWS 
    # -------------------------------------------------------- 
 
    before = len( 
        result 
    ) 
 
    result = ( 
        result 
        .dropna( 
            how="all" 
        ) 
        .reset_index( 
            drop=True 
        ) 
    ) 
 
    removed = ( 
        before - len(result) 
    ) 
 
    if removed > 0: 
 
        log.append( 
            f"Removed {removed} " 
            f"completely blank row(s)." 
        ) 
 
    # -------------------------------------------------------- 
    # 3. REMOVE BLANK COLUMNS 
    # -------------------------------------------------------- 
 
    blank_columns = [ 
        col 
        for col 
        in result.columns 
        if result[col].isna().all() 
    ] 
 
    if blank_columns: 
 
        result = result.drop( 
            columns=blank_columns 
        ) 
 
        log.append( 
            f"Removed {len(blank_columns)} " 
            f"completely blank column(s)." 
        ) 
 
    # -------------------------------------------------------- 
    # 4. TRIM TEXT 
    # -------------------------------------------------------- 
 
    whitespace_changes = 0 
 
    for col in result.columns: 
 
        if ( 
            result[col].dtype == "object" 
            or 
            pd.api.types.is_string_dtype( 
                result[col] 
            ) 
        ): 
 
            original_values = ( 
                result[col].copy() 
            ) 
 
            result[col] = ( 
                result[col] 
                .apply( 
                    lambda x: 
                    clean_text(x) 
                    if not pd.isna(x) 
                    else x 
                ) 
            ) 
 
            whitespace_changes += int( 
                ( 
                    original_values.astype(str) 
                    != 
                    result[col].astype(str) 
                ).sum() 
            ) 
 
    if whitespace_changes > 0: 
 
        log.append( 
            f"Trimmed/cleaned text values " 
            f"in {whitespace_changes} cell(s)." 
        ) 
 
    # -------------------------------------------------------- 
    # 5. REMOVE DUPLICATES 
    # -------------------------------------------------------- 
 
    before = len( 
        result 
    ) 
 
    result = ( 
        result 
        .drop_duplicates( 
            keep="first" 
        ) 
        .reset_index( 
            drop=True 
        ) 
    ) 
 
    removed = ( 
        before - len(result) 
    ) 
 
    if removed > 0: 
 
        log.append( 
            f"Removed {removed} " 
            f"duplicate row(s)." 
        ) 
 
    # -------------------------------------------------------- 
    # 6. NUMERIC CONVERSION 
    # -------------------------------------------------------- 
 
    for col in result.columns: 
 
        if result[col].dtype == "object": 
 
            original = result[col] 
 
            converted = pd.to_numeric( 
                original, 
                errors="coerce" 
            ) 
 
            non_null_original = ( 
                original.notna().sum() 
            ) 
 
            if non_null_original == 0: 
                continue 
 
            valid_numeric = ( 
                converted.notna().sum() 
            ) 
 
            ratio = ( 
                valid_numeric 
                / 
                non_null_original 
            ) 
 
            if ratio >= 0.80: 
 
                result[col] = ( 
                    converted 
                ) 
 
                log.append( 
                    f"Converted '{col}' " 
                    f"to numeric data type." 
                ) 
 
    # -------------------------------------------------------- 
    # 7. DATE CONVERSION 
    # -------------------------------------------------------- 
 
    for col in result.columns: 
 
        col_lower = str( 
            col 
        ).lower() 
 
        if ( 
            "date" in col_lower 
            or 
            "time" in col_lower 
        ): 
 
            if pd.api.types.is_datetime64_any_dtype( 
                result[col] 
            ): 
 
                continue 
 
            converted = pd.to_datetime( 
                result[col], 
                errors="coerce" 
            ) 
 
            original_non_null = ( 
                result[col].notna().sum() 
            ) 
 
            if original_non_null == 0: 
                continue 
 
            valid_dates = ( 
                converted.notna().sum() 
            ) 
 
            ratio = ( 
                valid_dates 
                / 
                original_non_null 
            ) 
 
            if ratio >= 0.70: 
 
                result[col] = ( 
                    converted 
                ) 
 
                log.append( 
                    f"Converted '{col}' " 
                    f"to date/time format." 
                ) 
 
    # -------------------------------------------------------- 
    # 8. HANDLE MISSING VALUES 
    # -------------------------------------------------------- 
 
    for col in result.columns: 
 
        missing_count = int( 
            result[col].isna().sum() 
        ) 
 
        if missing_count == 0: 
            continue 
 
        # Text 
        if ( 
            result[col].dtype == "object" 
            or 
            pd.api.types.is_string_dtype( 
                result[col] 
            ) 
        ): 
 
            result[col] = ( 
                result[col] 
                .fillna( 
                    "Unknown" 
                ) 
            ) 
 
            log.append( 
                f"Filled {missing_count} " 
                f"missing value(s) in " 
                f"'{col}' with 'Unknown'." 
            ) 
 
        # Numeric 
        elif pd.api.types.is_numeric_dtype( 
            result[col] 
        ): 
 
            median_value = ( 
                result[col].median() 
            ) 
 
            if pd.notna( 
                median_value 
            ): 
 
                result[col] = ( 
                    result[col] 
                    .fillna( 
                        median_value 
                    ) 
                ) 
 
                log.append( 
                    f"Filled {missing_count} " 
                    f"missing value(s) in " 
                    f"'{col}' with median " 
                    f"{median_value}." 
                ) 
 
        # Date 
        elif pd.api.types.is_datetime64_any_dtype( 
            result[col] 
        ): 
 
            mode = ( 
                result[col] 
                .mode() 
            ) 
 
            if not mode.empty: 
 
                result[col] = ( 
                    result[col] 
                    .fillna( 
                        mode.iloc[0] 
                    ) 
                ) 
 
                log.append( 
                    f"Filled {missing_count} " 
                    f"missing date value(s) " 
                    f"in '{col}' with the " 
                    f"most common date." 
                ) 
 
    # -------------------------------------------------------- 
    # FINAL LOG 
    # -------------------------------------------------------- 
 
    final_rows = len( 
        result 
    ) 
 
    if original_rows != final_rows: 
 
        log.append( 
            f"Rows changed from " 
            f"{original_rows} to " 
            f"{final_rows}." 
        ) 
 
    if not log: 
 
        log.append( 
            "No automatic cleaning " 
            "changes were required." 
        ) 
 
    return result, log 
 
 
# ============================================================ 
# MANUAL CLEANING 
# ============================================================ 
 
def apply_manual_cleaning( 
    df, 
    operation, 
    column=None, 
    value=None, 
    replacement=None 
): 
 
    result = df.copy() 
 
    log = [] 
 
    # -------------------------------------------------------- 
    # TRIM 
    # -------------------------------------------------------- 
 
    if operation == "TRIM": 
 
        if column and column in result.columns: 
 
            result[column] = ( 
                result[column] 
                .apply( 
                    lambda x: 
                    clean_text(x) 
                    if not pd.isna(x) 
                    else x 
                ) 
            ) 
 
            log.append( 
                f"Trimmed whitespace " 
                f"in '{column}'." 
            ) 
 
        else: 
 
            for col in result.columns: 
 
                if ( 
                    result[col].dtype == "object" 
                    or 
                    pd.api.types.is_string_dtype( 
                        result[col] 
                    ) 
                ): 
 
                    result[col] = ( 
                        result[col] 
                        .apply( 
                            lambda x: 
                            clean_text(x) 
                            if not pd.isna(x) 
                            else x 
                        ) 
                    ) 
 
            log.append( 
                "Trimmed whitespace in " 
                "all text columns." 
            ) 
 
    # -------------------------------------------------------- 
    # UPPER 
    # -------------------------------------------------------- 
 
    elif operation == "UPPER": 
 
        if column in result.columns: 
 
            result[column] = ( 
                result[column] 
                .apply( 
                    lambda x: 
                    str(x).upper() 
                    if not pd.isna(x) 
                    else x 
                ) 
            ) 
 
            log.append( 
                f"Converted '{column}' " 
                f"to uppercase." 
            ) 
 
    # -------------------------------------------------------- 
    # LOWER 
    # -------------------------------------------------------- 
 
    elif operation == "LOWER": 
 
        if column in result.columns: 
 
            result[column] = ( 
                result[column] 
                .apply( 
                    lambda x: 
                    str(x).lower() 
                    if not pd.isna(x) 
                    else x 
                ) 
            ) 
 
            log.append( 
                f"Converted '{column}' " 
                f"to lowercase." 
            ) 
 
    # -------------------------------------------------------- 
    # PROPER 
    # -------------------------------------------------------- 
 
    elif operation == "PROPER": 
 
        if column in result.columns: 
 
            result[column] = ( 
                result[column] 
                .apply( 
                    lambda x: 
                    str(x).title() 
                    if not pd.isna(x) 
                    else x 
                ) 
            ) 
 
            log.append( 
                f"Converted '{column}' " 
                f"to proper case." 
            ) 
 
    # -------------------------------------------------------- 
    # CLEAN 
    # -------------------------------------------------------- 
 
    elif operation == "CLEAN": 
 
        if column in result.columns: 
 
            result[column] = ( 
                result[column] 
                .apply( 
                    lambda x: 
                    clean_text(x) 
                    if not pd.isna(x) 
                    else x 
                ) 
            ) 
 
            log.append( 
                f"Cleaned text in " 
                f"'{column}'." 
            ) 
 
    # -------------------------------------------------------- 
    # REMOVE DUPLICATES 
    # -------------------------------------------------------- 
 
    elif operation == "REMOVE DUPLICATES": 
 
        before = len( 
            result 
        ) 
 
        result = ( 
            result 
            .drop_duplicates( 
                keep="first" 
            ) 
            .reset_index( 
                drop=True 
            ) 
        ) 
 
        removed = ( 
            before - len(result) 
        ) 
 
        log.append( 
            f"Removed {removed} " 
            f"duplicate row(s)." 
        ) 
 
    # -------------------------------------------------------- 
    # REMOVE BLANK ROWS 
    # -------------------------------------------------------- 
 
    elif operation == "REMOVE BLANK ROWS": 
 
        before = len( 
            result 
        ) 
 
        result = ( 
            result 
            .dropna( 
                how="all" 
            ) 
            .reset_index( 
                drop=True 
            ) 
        ) 
 
        removed = ( 
            before - len(result) 
        ) 
 
        log.append( 
            f"Removed {removed} " 
            f"completely blank row(s)." 
        ) 
 
    # -------------------------------------------------------- 
    # FILL MISSING VALUES 
    # -------------------------------------------------------- 
 
    elif operation == "FILL MISSING VALUES": 
 
        if column in result.columns: 
 
            fill_value = ( 
                value 
                if value is not None 
                else "Unknown" 
            ) 
 
            result[column] = ( 
                result[column] 
                .fillna( 
                    fill_value 
                ) 
            ) 
 
            log.append( 
                f"Filled missing values " 
                f"in '{column}' with " 
                f"'{fill_value}'." 
            ) 
 
    # -------------------------------------------------------- 
    # REPLACE VALUES 
    # -------------------------------------------------------- 
 
    elif operation == "REPLACE VALUES": 
 
        if column in result.columns: 
 
            result[column] = ( 
                result[column] 
                .replace( 
                    value, 
                    replacement 
                ) 
            ) 
 
            log.append( 
                f"Replaced '{value}' " 
                f"with '{replacement}' " 
                f"in '{column}'." 
            ) 
 
    # -------------------------------------------------------- 
    # CONVERT TO NUMBER 
    # -------------------------------------------------------- 
 
    elif operation == "CONVERT TO NUMBER": 
 
        if column in result.columns: 
 
            result[column] = ( 
                pd.to_numeric( 
                    result[column], 
                    errors="coerce" 
                ) 
            ) 
 
            log.append( 
                f"Converted '{column}' " 
                f"to numeric." 
            ) 
 
    # -------------------------------------------------------- 
    # CONVERT TO DATE 
    # -------------------------------------------------------- 
 
    elif operation == "CONVERT TO DATE": 
 
        if column in result.columns: 
 
            result[column] = ( 
                pd.to_datetime( 
                    result[column], 
                    errors="coerce" 
                ) 
            ) 
 
            log.append( 
                f"Converted '{column}' " 
                f"to date." 
            ) 
 
    return result, log 
 
 
# ============================================================ 
# GENERIC CLEANING VALIDATION 
# ============================================================ 
 
def count_whitespace_issues(df): 
 
    count = 0 
 
    for col in df.columns: 
 
        if ( 
            df[col].dtype == "object" 
            or 
            pd.api.types.is_string_dtype( 
                df[col] 
            ) 
        ): 
 
            values = ( 
                df[col] 
                .dropna() 
                .astype(str) 
            ) 
 
            count += int( 
                values.ne( 
                    values.str.strip() 
                ).sum() 
            ) 
 
    return count 
 
 
def validate_cleaning( 
    raw_df, 
    cleaned_df 
): 
 
    checks = [] 
 
    # -------------------------------------------------------- 
    # COUNTS 
    # -------------------------------------------------------- 
 
    raw_rows = len( 
        raw_df 
    ) 
 
    clean_rows = len( 
        cleaned_df 
    ) 
 
    raw_cols = len( 
        raw_df.columns 
    ) 
 
    clean_cols = len( 
        cleaned_df.columns 
    ) 
 
    # -------------------------------------------------------- 
    # ROW COUNT 
    # -------------------------------------------------------- 
 
    if clean_rows <= raw_rows: 
 
        status = "PASS" 
 
        details = ( 
            "Rows were preserved or " 
            "removed during cleaning." 
        ) 
 
    else: 
 
        status = "WARNING" 
 
        details = ( 
            "Cleaned data contains more " 
            "rows than the raw data." 
        ) 
 
    checks.append({ 
        "Check": "Row count", 
        "Before": raw_rows, 
        "After": clean_rows, 
        "Status": status, 
        "Details": details 
    }) 
 
    # -------------------------------------------------------- 
    # COLUMN COUNT 
    # -------------------------------------------------------- 
 
    if clean_cols == raw_cols: 
 
        status = "PASS" 
 
        details = ( 
            "Column count is unchanged." 
        ) 
 
    else: 
 
        status = "INFO" 
 
        details = ( 
            "Column count changed " 
            "during cleaning." 
        ) 
 
    checks.append({ 
        "Check": "Column count", 
        "Before": raw_cols, 
        "After": clean_cols, 
        "Status": status, 
        "Details": details 
    }) 
 
    # -------------------------------------------------------- 
    # BLANK ROWS 
    # -------------------------------------------------------- 
 
    raw_blank_rows = int( 
        raw_df.isna() 
        .all(axis=1) 
        .sum() 
    ) 
 
    clean_blank_rows = int( 
        cleaned_df.isna() 
        .all(axis=1) 
        .sum() 
    ) 
 
    if clean_blank_rows == 0: 
 
        status = "PASS" 
 
        details = ( 
            "No completely blank " 
            "rows remain." 
        ) 
 
    elif clean_blank_rows < raw_blank_rows: 
 
        status = "PASS" 
 
        details = ( 
            f"Blank rows reduced " 
            f"from {raw_blank_rows} " 
            f"to {clean_blank_rows}." 
        ) 
 
    else: 
 
        status = "WARNING" 
 
        details = ( 
            "Completely blank rows " 
            "still exist." 
        ) 
 
    checks.append({ 
        "Check": "Completely blank rows", 
        "Before": raw_blank_rows, 
        "After": clean_blank_rows, 
        "Status": status, 
        "Details": details 
    }) 
 
    # -------------------------------------------------------- 
    # DUPLICATES 
    # -------------------------------------------------------- 
 
    raw_duplicates = int( 
        raw_df.duplicated().sum() 
    ) 
 
    clean_duplicates = int( 
        cleaned_df.duplicated().sum() 
    ) 
 
    if clean_duplicates == 0: 
 
        status = "PASS" 
 
        details = ( 
            "No duplicate rows remain." 
        ) 
 
    elif clean_duplicates < raw_duplicates: 
 
        status = "PASS" 
 
        details = ( 
            f"Duplicate rows reduced " 
            f"from {raw_duplicates} " 
            f"to {clean_duplicates}." 
        ) 
 
    else: 
 
        status = "INFO" 
 
        details = ( 
            "Duplicate rows were " 
            "not removed." 
        ) 
 
    checks.append({ 
        "Check": "Duplicate rows", 
        "Before": raw_duplicates, 
        "After": clean_duplicates, 
        "Status": status, 
        "Details": details 
    }) 
 
    # -------------------------------------------------------- 
    # WHITESPACE 
    # -------------------------------------------------------- 
 
    raw_whitespace = ( 
        count_whitespace_issues( 
            raw_df 
        ) 
    ) 
 
    clean_whitespace = ( 
        count_whitespace_issues( 
            cleaned_df 
        ) 
    ) 
 
    if clean_whitespace == 0: 
 
        status = "PASS" 
 
        details = ( 
            "No leading or trailing " 
            "whitespace remains." 
        ) 
 
    elif clean_whitespace < raw_whitespace: 
 
        status = "PASS" 
 
        details = ( 
            f"Whitespace issues reduced " 
            f"from {raw_whitespace} " 
            f"to {clean_whitespace}." 
        ) 
 
    else: 
 
        status = "INFO" 
 
        details = ( 
            "Whitespace issues were " 
            "not completely removed." 
        ) 
 
    checks.append({ 
        "Check": "Leading/trailing whitespace", 
        "Before": raw_whitespace, 
        "After": clean_whitespace, 
        "Status": status, 
        "Details": details 
    }) 
 
    # -------------------------------------------------------- 
    # MISSING VALUES 
    # -------------------------------------------------------- 
 
    raw_missing = int( 
        raw_df.isna() 
        .sum() 
        .sum() 
    ) 
 
    clean_missing = int( 
        cleaned_df.isna() 
        .sum() 
        .sum() 
    ) 
 
    if clean_missing < raw_missing: 
 
        status = "PASS" 
 
        details = ( 
            f"Missing values reduced " 
            f"from {raw_missing} " 
            f"to {clean_missing}." 
        ) 
 
    elif clean_missing == raw_missing: 
 
        status = "INFO" 
 
        details = ( 
            "Missing-value count " 
            "is unchanged." 
        ) 
 
    else: 
 
        status = "WARNING" 
 
        details = ( 
            "Missing-value count " 
            "increased." 
        ) 
 
    checks.append({ 
        "Check": "Missing values", 
        "Before": raw_missing, 
        "After": clean_missing, 
        "Status": status, 
        "Details": details 
    }) 
 
    # -------------------------------------------------------- 
    # DATA TYPES 
    # -------------------------------------------------------- 
 
    common_columns = [ 
        col 
        for col in raw_df.columns 
        if col in cleaned_df.columns 
    ] 
 
    dtype_changes = 0 
 
    for col in common_columns: 
 
        raw_type = str( 
            raw_df[col].dtype 
        ) 
 
        clean_type = str( 
            cleaned_df[col].dtype 
        ) 
 
        if raw_type != clean_type: 
 
            dtype_changes += 1 
 
    if dtype_changes > 0: 
 
        status = "PASS" 
 
        details = ( 
            f"{dtype_changes} column(s) " 
            f"had data type changes." 
        ) 
 
    else: 
 
        status = "INFO" 
 
        details = ( 
            "No data type changes " 
            "were detected." 
        ) 
 
    checks.append({ 
        "Check": "Data types", 
        "Before": "Original", 
        "After": "Cleaned", 
        "Status": status, 
        "Details": details 
    }) 
 
    # -------------------------------------------------------- 
    # COLUMN NAMES 
    # -------------------------------------------------------- 
 
    if list(raw_df.columns) == list( 
        cleaned_df.columns 
    ): 
 
        status = "PASS" 
 
        details = ( 
            "Column names are consistent." 
        ) 
 
    else: 
 
        status = "INFO" 
 
        details = ( 
            "Column names were changed." 
        ) 
 
    checks.append({ 
        "Check": "Column names", 
        "Before": ", ".join( 
            map( 
                str, 
                raw_df.columns 
            ) 
        ), 
        "After": ", ".join( 
            map( 
                str, 
                cleaned_df.columns 
            ) 
        ), 
        "Status": status, 
        "Details": details 
    }) 
 
    # -------------------------------------------------------- 
    # OVERALL STATUS 
    # -------------------------------------------------------- 
 
    warning_count = sum( 
        1 
        for item in checks 
        if item["Status"] == "WARNING" 
    ) 
 
    if warning_count == 0: 
 
        overall_status = "PASS" 
 
    else: 
 
        overall_status = "WARNING" 
 
    return ( 
        pd.DataFrame( 
            checks 
        ), 
        overall_status 
    ) 
 
 
# ============================================================ 
# CALCULATED COLUMN ENGINE 
# ============================================================ 
 
def create_calculated_column(df, request):
    """
    Creates a calculated column using arithmetic expressions.
    Supports +, -, *, /, parentheses, numeric constants, and
    existing dataframe columns, e.g.:
    Total With Tax = Bill_amt + (Bill_amt * 0.05)
    """
    if df is None:
        return None, None, "No dataset available."

    text = extract_create_calculated_column_request(request)
    if not text or "=" not in text:
        return None, None, (
            "Please use: Create a calculated column "
            "Revenue = Quantity * Price"
        )

    new_column, formula = text.split("=", 1)
    new_column = new_column.strip().strip("`\"'")
    formula = formula.strip()

    if not new_column:
        return None, None, "Calculated column name is missing."
    if not formula:
        return None, None, "The calculation formula is missing."

    existing_column = find_column(df, new_column)
    if existing_column is not None:
        if is_calculated_column(existing_column):
            return None, None, (
                f"Calculated column '{existing_column}' already exists. "
                "Delete it first before creating it again."
            )
        return None, None, (
            f"'{existing_column}' is already an original Excel column. "
            "Calculated columns cannot overwrite original columns."
        )

    # Parse only arithmetic expressions; never use eval().
    try:
        tree = ast.parse(formula, mode="eval")
    except SyntaxError as e:
        return None, None, f"Invalid formula syntax: {e.msg}"

    def evaluate(node):
        if isinstance(node, ast.Expression):
            return evaluate(node.body)
        if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
            return float(node.value)
        if isinstance(node, ast.Name):
            column = find_column(df, node.id)
            if column is None:
                raise ValueError(f"Column '{node.id}' was not found.")
            return df[column]
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.UAdd, ast.USub)):
            value = evaluate(node.operand)
            return value if isinstance(node.op, ast.UAdd) else -value
        if isinstance(node, ast.BinOp) and isinstance(
            node.op, (ast.Add, ast.Sub, ast.Mult, ast.Div)
        ):
            left = evaluate(node.left)
            right = evaluate(node.right)
            if isinstance(node.op, ast.Add):
                return left + right
            if isinstance(node.op, ast.Sub):
                return left - right
            if isinstance(node.op, ast.Mult):
                return left * right
            return left / right
        raise ValueError(
            "Only column names, numbers, +, -, *, /, and parentheses are supported."
        )

    try:
        value = evaluate(tree)
        result = df.copy()
        result[new_column] = value
        result[new_column] = result[new_column].replace(
            [np.inf, -np.inf], np.nan
        )
    except Exception as e:
        return None, None, f"Unable to calculate column: {e}"

    if not is_calculated_column(new_column):
        st.session_state.calculated_columns.append(new_column)

    return result, new_column, f"Created calculated column '{new_column}'."


# ============================================================ 
# GENERIC KPI FUNCTIONS 
# ============================================================ 
 
def total_rows(df): 
 
    return len(df) 
 
 
def numeric_summary(df): 
 
    numeric = df.select_dtypes( 
        include=np.number 
    ) 
 
    if numeric.empty: 
 
        return pd.DataFrame() 
 
    return ( 
        numeric 
        .describe() 
        .T 
        .reset_index() 
        .rename( 
            columns={ 
                "index": "Column" 
            } 
        ) 
    ) 
 
 
def categorical_summary(df): 
 
    rows = [] 
 
    for col in df.columns: 
 
        if ( 
            df[col].dtype == "object" 
            or 
            pd.api.types.is_string_dtype( 
                df[col] 
            ) 
        ): 
 
            rows.append({ 
                "Column": col, 
 
                "Unique Values": 
                    df[col].nunique( 
                        dropna=True 
                    ), 
 
                "Missing": 
                    int( 
                        df[col].isna().sum() 
                    ) 
            }) 
 
    return pd.DataFrame( 
        rows 
    ) 
 
 
# ============================================================ 
# TOP VALUES 
# ============================================================ 
 
def top_categories( 
    df, 
    column, 
    n=5 
): 
 
    if column not in df.columns: 
 
        return pd.DataFrame() 
 
    result = ( 
        df[column] 
        .value_counts() 
        .head(n) 
        .reset_index() 
    ) 
 
    result.columns = [ 
        "Category", 
        "Count" 
    ] 
 
    return result 
 
 
# ============================================================ 
# GENERIC REVENUE / NUMERIC ANALYSIS 
# ============================================================ 
 
def revenue_analysis(df): 
 
    numeric_cols = ( 
        find_numeric_columns( 
            df 
        ) 
    ) 
 
    if not numeric_cols: 
 
        return None 
 
    preferred = [ 
        col 
        for col in numeric_cols 
        if any( 
            word in str(col).lower() 
            for word in [ 
                "revenue", 
                "sales", 
                "amount", 
                "price", 
                "total", 
                "bill" 
            ] 
        ) 
    ] 
 
    if preferred: 
 
        column = preferred[0] 
 
    else: 
 
        column = numeric_cols[0] 
 
    return { 
        "column": column, 
 
        "sum": df[column].sum(), 
 
        "average": df[column].mean(), 
 
        "minimum": df[column].min(), 
 
        "maximum": df[column].max() 
    } 
 
 
# ============================================================ 
# GROUP ANALYSIS 
# ============================================================ 
 
def group_numeric_by_category(
    df,
    category_column,
    numeric_column,
    n=None
):
    if category_column not in df.columns or numeric_column not in df.columns:
        return pd.DataFrame()
    grouped = df.groupby(category_column, dropna=False)[numeric_column].sum().sort_values(ascending=False)
    if n is not None:
        grouped = grouped.head(n)
    return grouped.reset_index()


def daily_numeric_trend( 
    df, 
    date_column, 
    numeric_column 
): 
 
    if ( 
        date_column not in df.columns 
        or 
        numeric_column not in df.columns 
    ): 
 
        return pd.DataFrame() 
 
    temp = df[ 
        [ 
            date_column, 
            numeric_column 
        ] 
    ].copy() 
 
    temp[date_column] = ( 
        pd.to_datetime( 
            temp[date_column], 
            errors="coerce" 
        ) 
    ) 
 
    temp = temp.dropna( 
        subset=[ 
            date_column 
        ] 
    ) 
 
    result = ( 
        temp.groupby( 
            temp[date_column].dt.date 
        )[numeric_column] 
        .sum() 
        .reset_index() 
    ) 
 
    return result 
 
 
# ============================================================ 
# PIVOT TABLE 
# ============================================================ 
 
def create_pivot_table( 
    df 
): 
 
    if df is None or df.empty: 
 
        return pd.DataFrame() 
 
    categorical = [ 
        col 
        for col in df.columns 
        if ( 
            df[col].dtype == "object" 
            or 
            pd.api.types.is_string_dtype( 
                df[col] 
            ) 
        ) 
    ] 
 
    numeric = ( 
        find_numeric_columns( 
            df 
        ) 
    ) 
 
    if not categorical or not numeric: 
 
        return pd.DataFrame() 
 
    category_column = ( 
        categorical[0] 
    ) 
 
    numeric_column = ( 
        numeric[0] 
    ) 
 
    pivot = pd.pivot_table( 
        df, 
        index=category_column, 
        values=numeric_column, 
        aggfunc=[ 
            "count", 
            "sum", 
            "mean", 
            "min", 
            "max" 
        ], 
        fill_value=0 
    ) 
 
    pivot = pivot.reset_index() 
 
    pivot.columns = [ 
        "_".join( 
            [ 
                str(x) 
                for x in col 
                if str(x) != "" 
            ] 
        ).strip("_") 
        if isinstance( 
            col, 
            tuple 
        ) 
        else str(col) 
        for col in pivot.columns 
    ] 
 
    return pivot 
 
 
# ============================================================
# VISUALIZATION HELPERS
# ============================================================

def normalize_for_match(value):
    """Normalize text while preserving both spaced and compact forms."""
    return re.sub(r"[^a-z0-9]+", " ", str(value).lower()).strip()


def compact_for_match(value):
    """Return a compact alphanumeric form so Food Item == Food_Item == FoodItem."""
    return re.sub(r"[^a-z0-9]+", "", str(value).lower())


def resolve_requested_column(df, request, candidates, preferred_words=None, phrase=None):
    """Resolve the column the user actually named, with phrase-aware matching."""
    if df is None or not candidates:
        return None

    request_norm = normalize_for_match(request)
    request_tokens = set(request_norm.split())
    preferred_words = [normalize_for_match(w) for w in (preferred_words or [])]
    preferred_words = [w for w in preferred_words if w]

    # 1. If the caller identified a phrase (for example the text after
    #    "by"), resolve that phrase first. This prevents fallback columns
    #    such as Hotel Name from winning over an explicitly requested
    #    Food Item column.
    if phrase:
        phrase_norm = normalize_for_match(phrase)
        phrase_compact = compact_for_match(phrase)
        phrase_tokens = set(phrase_norm.split())
        if phrase_norm:
            exact_phrase = []
            for col in candidates:
                col_norm = normalize_for_match(col)
                col_compact = compact_for_match(col)
                # Treat spaces, underscores and camel/concatenated names as
                # equivalent: Food Item == Food_Item == FoodItem.
                if (
                    col_norm == phrase_norm
                    or phrase_compact == col_compact
                    or (phrase_compact and phrase_compact in col_compact)
                    or (col_compact and col_compact in phrase_compact)
                ):
                    exact_phrase.append(col)
            if exact_phrase:
                return max(exact_phrase, key=lambda c: len(compact_for_match(c)))

            phrase_scored = []
            def _stem_token(token):
                token = str(token).lower()
                if len(token) > 4 and token.endswith("ies"):
                    return token[:-3] + "y"
                if len(token) > 3 and token.endswith("s"):
                    return token[:-1]
                return token

            phrase_stems = {_stem_token(t) for t in phrase_tokens}
            for col in candidates:
                col_norm = normalize_for_match(col)
                col_compact = compact_for_match(col)
                col_tokens = set(col_norm.split())
                col_stems = {_stem_token(t) for t in col_tokens}
                overlap = len(col_stems & phrase_stems)
                compact_bonus = 1 if phrase_compact and phrase_compact in col_compact else 0
                # An explicit phrase must have a real match. Compact matching
                # handles concatenated Excel headers such as FoodItemName.
                if overlap or compact_bonus:
                    coverage = overlap / max(len(phrase_tokens), 1)
                    phrase_scored.append((compact_bonus, coverage, overlap, -len(col_tokens), col))
            if phrase_scored:
                phrase_scored.sort(key=lambda x: x[:-1], reverse=True)
                return phrase_scored[0][-1]

    # 2. Exact normalized column-name match anywhere in the request.
    exact_matches = []
    request_compact = compact_for_match(request)
    for col in candidates:
        col_norm = normalize_for_match(col)
        col_compact = compact_for_match(col)
        if (
            (col_norm and col_norm in request_norm)
            or (col_compact and col_compact in request_compact)
        ):
            exact_matches.append(col)
    if exact_matches:
        return max(exact_matches, key=lambda c: len(compact_for_match(c)))

    # 3. Strong token overlap. Prefer the column with the greatest
    #    coverage of the user's words rather than an arbitrary semantic
    #    preference such as "hotel".
    scored = []
    for col in candidates:
        col_norm = normalize_for_match(col)
        col_tokens = set(col_norm.split())
        def _stem_token(token):
            token = str(token).lower()
            if len(token) > 4 and token.endswith("ies"):
                return token[:-3] + "y"
            if len(token) > 3 and token.endswith("s"):
                return token[:-1]
            return token
        overlap = len({_stem_token(t) for t in col_tokens} & {_stem_token(t) for t in request_tokens})
        if overlap:
            coverage = overlap / max(len(col_tokens), 1)
            request_coverage = overlap / max(len(request_tokens), 1)
            preferred_bonus = sum(1 for word in preferred_words if word in col_norm)
            scored.append((overlap, coverage, request_coverage, preferred_bonus, len(col_norm), col))

    if scored:
        scored.sort(key=lambda item: item[:-1], reverse=True)
        return scored[0][-1]

    # 4. Semantic fallback only when the user did not explicitly identify
    #    a usable column.
    for word in preferred_words:
        for col in candidates:
            if word in normalize_for_match(col):
                return col

    return None


def _extract_by_phrase(request):
    """Return the phrase following 'by/against/for each/per' when present."""
    raw = str(request or "").strip()
    patterns = [
        r"\bby\s+(.+)$",
        r"\bagainst\s+(.+)$",
        r"\bfor\s+each\s+(.+)$",
        r"\bper\s+(.+)$",
        r"\bitem[- ]?wise\s+(.+)$",
    ]
    for pattern in patterns:
        match = re.search(pattern, raw, flags=re.IGNORECASE)
        if match:
            phrase = match.group(1).strip()
            phrase = re.sub(
                r"\b(?:as\s+)?(?:a|an|the)?\s*(?:bar|line|pie|donut|scatter|histogram|chart|graph|plot)\b.*$",
                "",
                phrase,
                flags=re.IGNORECASE,
            ).strip()
            return phrase
    return None


def _extract_metric_phrase(request):
    """Return the metric phrase before 'by/against/for each/per'."""
    raw = str(request or "").strip()
    match = re.search(
        r"\b(?:by|against|for\s+each|per)\b",
        raw,
        flags=re.IGNORECASE,
    )
    if not match:
        return None

    before = raw[:match.start()]
    # Remove common chart-command words so they cannot affect metric matching.
    before = re.sub(
        r"\b(?:show|create|make|display|give|draw|plot|visuali[sz]e|a|an|the|bar|line|pie|donut|scatter|histogram|chart|graph|plot|horizontal|vertical|stacked)\b",
        " ",
        before,
        flags=re.IGNORECASE,
    )
    before = re.sub(r"\bof\b", " ", before, flags=re.IGNORECASE)
    before = re.sub(r"\s+", " ", before).strip()
    return before or None


def _extract_top_n(request):
    """Extract an explicit Top-N limit such as 'top 3', 'top 10', or 'top three'."""
    raw = str(request or "").strip().lower()
    match = re.search(r"\btop\s*[- ]?\s*(\d+)\b", raw)
    if match:
        try:
            return max(1, int(match.group(1)))
        except (TypeError, ValueError):
            pass

    number_words = {
        "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
        "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
        "eleven": 11, "twelve": 12, "thirteen": 13, "fourteen": 14,
        "fifteen": 15, "twenty": 20,
    }
    match = re.search(r"\btop\s+(" + "|".join(number_words.keys()) + r")\b", raw)
    if match:
        return number_words[match.group(1)]

    return None


def _extract_vs_phrases(request):
    raw = str(request or "").strip()
    match = re.search(r"(.+?)\s+(?:vs\.?|versus|against)\s+(.+)$", raw, flags=re.IGNORECASE)
    if not match:
        return None, None
    left = re.sub(r"\b(?:show|create|make|display|give|draw|plot|a|an|the|scatter|scatter\s+plot|chart|graph)\b", " ", match.group(1), flags=re.IGNORECASE)
    right = re.sub(r"\b(?:chart|graph|plot)\b", " ", match.group(2), flags=re.IGNORECASE)
    return left.strip(), right.strip()


def _find_status_column(df, requested_values=None):
    requested_values = [str(v).lower() for v in (requested_values or [])]
    candidates = []
    for col in df.columns:
        if not (df[col].dtype == "object" or pd.api.types.is_string_dtype(df[col])):
            continue
        name = normalize_for_match(col)
        values = set(df[col].dropna().astype(str).str.strip().str.lower().unique())
        hits = sum(1 for v in requested_values if v in values)
        name_bonus = 2 if any(w in name for w in ["status", "state", "delivery status", "order status"]) else 0
        candidates.append((hits, name_bonus, -df[col].nunique(dropna=True), col))
    if not candidates:
        return None
    candidates.sort(reverse=True)
    best = candidates[0]
    return best[-1] if best[0] > 0 or best[1] > 0 else None


def _requested_status_values(request):
    raw = str(request or "").lower()
    known = ["delivered", "cancelled", "canceled", "pending", "processing", "completed", "failed", "returned", "refunded", "shipped", "out for delivery"]
    return [v for v in known if re.search(r"\b" + re.escape(v) + r"\b", raw)]


def _is_count_metric(request):
    text = normalize_for_match(request)
    return any(p in text for p in ["number of orders", "number of order", "orders", "order count", "count of orders", "count orders", "count", "number of records", "records", "frequency"])


def _is_quantity_metric(request):
    text = normalize_for_match(request)
    return any(w in text.split() for w in ["quantity", "qty", "units"])


def _extract_metric_column(df, request, numeric_cols, metric_phrase=None):
    if not numeric_cols:
        return None
    if metric_phrase:
        # When the user supplied an explicit metric phrase (for example
        # "delivery status"), require an actual metric word before accepting
        # a numeric column. This prevents Delivery_Partner_ID from being
        # selected merely because the phrase contains the word "delivery".
        metric_words = ["revenue", "sales", "amount", "profit", "price", "bill", "quantity", "qty", "value", "units"]
        metric_norm = normalize_for_match(metric_phrase)
        if not any(word in metric_norm.split() for word in metric_words):
            return None
        return resolve_requested_column(
            df, metric_phrase, numeric_cols,
            preferred_words=metric_words,
            phrase=metric_phrase,
        )
    return resolve_requested_column(
        df, request, numeric_cols,
        preferred_words=["revenue", "sales", "amount", "profit", "price", "bill", "quantity", "qty", "value"]
    )


# ============================================================
# ADVANCED VISUALIZATION ENGINE
# ============================================================


def _viz_tokens(value):
    return [t for t in normalize_for_match(value).split() if t]


def _viz_stem(token):
    token = str(token).lower()
    replacements = {
        "cities": "city", "countries": "country", "hotels": "hotel",
        "customers": "customer", "orders": "order", "items": "item",
        "products": "product", "sales": "sale", "revenues": "revenue",
        "quantities": "quantity", "units": "unit", "categories": "category",
        "departments": "department", "regions": "region", "partners": "partner",
    }
    return replacements.get(token, token[:-1] if len(token) > 4 and token.endswith("s") else token)


def _viz_clean_phrase(value):
    text = str(value or "")
    text = re.sub(
        r"\b(?:show|create|make|display|draw|plot|generate|give|visualize|visualise|a|an|the|chart|graph|plot|visualization|visualisation)\b",
        " ", text, flags=re.IGNORECASE
    )
    text = re.sub(
        r"\b(?:horizontal|vertical|stacked|clustered|grouped|top\s+\d+|top\s+(?:one|two|three|four|five|six|seven|eight|nine|ten))\b",
        " ", text, flags=re.IGNORECASE
    )
    return re.sub(r"\s+", " ", text).strip(" ,")


def _extract_top_n_viz(request):
    text = str(request or "")
    words = {
        "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
        "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
        "twenty": 20, "fifty": 50, "hundred": 100,
    }
    m = re.search(r"\btop\s+(\d+)\b", text, re.I)
    if m:
        return int(m.group(1))
    m = re.search(r"\btop\s+(one|two|three|four|five|six|seven|eight|nine|ten|twenty|fifty|hundred)\b", text, re.I)
    if m:
        return words[m.group(1).lower()]
    m = re.search(r"\b(?:first|highest|largest|best)\s+(\d+)\b", text, re.I)
    if m:
        return int(m.group(1))
    return None


def _extract_by_dimensions_viz(request):
    """Extract one or more dimension phrases after 'by', 'per', etc."""
    raw = str(request or "").strip()
    m = re.search(r"\b(?:by|per|for\s+each|against)\s+(.+)$", raw, re.I)
    if not m:
        return []
    tail = m.group(1).strip()
    tail = re.split(
        r"\b(?:where|with|using|as|colored\s+by|colour(?:ed)?\s+by|split\s+by|broken\s+down\s+by)\b",
        tail, maxsplit=1, flags=re.I
    )[0]
    tail = re.sub(
        r"\b(?:as\s+)?(?:a|an|the)?\s*(?:horizontal|vertical|stacked|clustered|grouped)?\s*(?:bar|column|line|area|pie|donut|doughnut|scatter|histogram|box|violin|heatmap|chart|graph|plot)\b.*$",
        "", tail, flags=re.I
    ).strip(" ,")
    if not tail:
        return []
    parts = re.split(r"\s*(?:,|\band\b|&|\bversus\b|\bvs\.?\b)\s*", tail, flags=re.I)
    return [p.strip() for p in parts if p.strip()]


def _extract_vs_viz(request):
    raw = str(request or "")
    m = re.search(r"(.+?)\s+(?:vs\.?|versus|against)\s+(.+)$", raw, re.I)
    if not m:
        return None, None
    return _viz_clean_phrase(m.group(1)), _viz_clean_phrase(m.group(2))


def _resolve_viz_column(df, phrase, candidates, strong=True):
    if df is None or not phrase or not candidates:
        return None
    phrase_norm = normalize_for_match(phrase)
    phrase_compact = compact_for_match(phrase)
    phrase_tokens = {_viz_stem(x) for x in _viz_tokens(phrase)}

    # Exact / normalized / compact matches first.
    exact = []
    for col in candidates:
        cn = normalize_for_match(col)
        cc = compact_for_match(col)
        if cn == phrase_norm or cc == phrase_compact or (phrase_compact and phrase_compact in cc) or (cc and cc in phrase_compact):
            exact.append(col)
    if exact:
        return max(exact, key=lambda c: len(compact_for_match(c)))

    scored = []
    for col in candidates:
        col_tokens = {_viz_stem(x) for x in _viz_tokens(col)}
        overlap = len(phrase_tokens & col_tokens)
        if overlap:
            coverage = overlap / max(len(phrase_tokens), 1)
            col_coverage = overlap / max(len(col_tokens), 1)
            scored.append((coverage, col_coverage, overlap, -len(col_tokens), col))
    if scored:
        scored.sort(key=lambda x: x[:-1], reverse=True)
        return scored[0][-1]
    return None


def _viz_dimension_candidates(df):
    numeric = find_numeric_columns(df)
    text = find_text_columns(df)
    dates = find_date_columns(df)
    dimensions = list(text)
    identifier_terms = {
        "id", "identifier", "code", "number", "no", "customer", "order",
        "hotel", "store", "branch", "account", "partner", "employee", "invoice",
        "transaction", "product"
    }
    for col in df.columns:
        if col in dimensions:
            continue
        norm = normalize_for_match(col)
        compact = compact_for_match(col)
        toks = set(norm.split())
        if toks & identifier_terms or any(k in compact for k in (
            "customerid", "orderid", "hotelid", "storeid", "branchid", "partnerid",
            "employeeid", "invoiceid", "transactionid", "productid"
        )):
            dimensions.append(col)
    for col in dates:
        if col not in dimensions:
            dimensions.append(col)
    return dimensions


def _viz_measure_candidates(df):
    return find_numeric_columns(df)


def _resolve_viz_measure(df, request, metric_phrase=None):
    numeric = _viz_measure_candidates(df)
    if not numeric:
        return None
    phrase = metric_phrase or ""
    if phrase:
        resolved = _resolve_viz_column(df, phrase, numeric)
        if resolved:
            return resolved
    text = normalize_for_match(request)
    # Explicit business metric vocabulary gets priority over arbitrary numeric columns.
    preferred_terms = [
        ("revenue", ["revenue", "sales", "sale_amount", "amount", "bill", "total"]),
        ("profit", ["profit", "margin"]),
        ("cost", ["cost", "expense", "spend"]),
        ("price", ["price", "rate", "fare"]),
        ("quantity", ["quantity", "qty", "units", "unit"]),
        ("discount", ["discount"]),
        ("tax", ["tax", "gst"]),
    ]
    for trigger, terms in preferred_terms:
        if trigger in text:
            for col in numeric:
                cn = normalize_for_match(col)
                if any(term in cn for term in terms):
                    return col
    # Never choose a numeric identifier as a measure unless the user explicitly asks for it.
    safe = [c for c in numeric if not re.search(r"(?:^|_)(?:id|code|number|no)$|(?:id|_id)$", str(c), re.I)]
    return safe[0] if safe else numeric[0]


def _viz_aggregation(request):
    text = normalize_for_match(request)
    if re.search(r"\b(?:average|avg|mean)\b", text):
        return "mean"
    if re.search(r"\b(?:median|midpoint)\b", text):
        return "median"
    if re.search(r"\b(?:minimum|min|lowest|smallest)\b", text):
        return "min"
    if re.search(r"\b(?:maximum|max|highest|largest)\b", text):
        return "max"
    if re.search(r"\b(?:count|number of|orders|records|rows|frequency|volume)\b", text) and not re.search(r"\b(?:revenue|sales|amount|price|profit|cost|quantity|units)\b", text):
        return "count"
    return "sum"


def _viz_chart_type(request, df):
    text = normalize_for_match(request)
    if any(x in text for x in ["heatmap", "heat map", "correlation matrix", "correlation heat"]):
        return "heatmap"
    if any(x in text for x in ["waterfall", "bridge chart"]):
        return "waterfall"
    if any(x in text for x in ["violin"]):
        return "violin"
    if any(x in text for x in ["box plot", "boxplot", "box chart"]):
        return "box"
    if any(x in text for x in ["histogram", "frequency distribution"]):
        return "histogram"
    if any(x in text for x in ["donut", "doughnut"]):
        return "donut"
    if any(x in text for x in ["pie", "share of", "share by", "proportion"]):
        return "pie"
    if any(x in text for x in ["scatter", "relationship", "correlation between", "plot x and y"]):
        return "scatter"
    if any(x in text for x in ["area chart", "area graph", "filled line"]):
        return "area"
    if any(x in text for x in ["line chart", "line graph", "trend", "over time", "timeline", "time series"]):
        return "time"
    if "stacked" in text:
        return "stacked_bar"
    if any(x in text for x in ["bar chart", "bar graph", "bars", "column chart", "column graph", "ranking", "ranked", "compare", "comparison", "top "]):
        return "bar"
    # Automatic fallback: choose the visualization that best fits the data.
    if find_date_columns(df) and find_numeric_columns(df):
        return "time"
    if _viz_dimension_candidates(df) and find_numeric_columns(df):
        return "bar"
    if len(find_numeric_columns(df)) >= 2:
        return "scatter"
    if len(find_numeric_columns(df)) == 1:
        return "histogram"
    return "count"


def _viz_explicit_metric_phrase(request):
    raw = str(request or "")
    m = re.search(r"\b(?:by|per|for\s+each|against)\b", raw, re.I)
    if not m:
        return None
    before = raw[:m.start()]
    before = _viz_clean_phrase(before)
    before = re.sub(r"\b(?:top\s+\d+|top\s+(?:one|two|three|four|five|six|seven|eight|nine|ten))\b", "", before, flags=re.I)
    return before.strip(" ,") or None


def _viz_build_grouped(df, category_col, value_col, aggregation="sum", top_n=None):
    if not category_col:
        return pd.DataFrame(), "Value"
    if aggregation == "count":
        result = df[category_col].fillna("Unknown").value_counts(dropna=False).rename("Count").reset_index()
        result.columns = [category_col, "Count"]
        value_name = "Count"
    else:
        temp = df[[category_col, value_col]].copy()
        temp[category_col] = temp[category_col].fillna("Unknown")
        temp[value_col] = pd.to_numeric(temp[value_col], errors="coerce")
        temp = temp.dropna(subset=[value_col])
        result = temp.groupby(category_col, dropna=False)[value_col].agg(aggregation).reset_index()
        value_name = str(value_col)
    result = result.sort_values(value_name, ascending=False, kind="stable")
    if top_n:
        result = result.head(top_n)
    return result, value_name


def _viz_prepare_time(df, date_col, value_col, aggregation, request):
    temp = df[[date_col] + ([value_col] if value_col else [])].copy()
    temp[date_col] = pd.to_datetime(temp[date_col], errors="coerce")
    temp = temp.dropna(subset=[date_col])
    text = normalize_for_match(request)
    if "hour" in text:
        period = temp[date_col].dt.floor("h")
        label = "Hourly"
    elif "week" in text:
        period = temp[date_col].dt.to_period("W").dt.start_time
        label = "Weekly"
    elif "month" in text:
        period = temp[date_col].dt.to_period("M").dt.start_time
        label = "Monthly"
    elif "quarter" in text:
        period = temp[date_col].dt.to_period("Q").dt.start_time
        label = "Quarterly"
    elif "year" in text or "annual" in text:
        period = temp[date_col].dt.to_period("Y").dt.start_time
        label = "Yearly"
    else:
        period = temp[date_col].dt.floor("D")
        label = "Daily"
    temp["__period__"] = period
    if aggregation == "count":
        out = temp.groupby("__period__").size().reset_index(name="Count")
        value_name = "Count"
    else:
        if not value_col:
            return pd.DataFrame(), None, label
        temp[value_col] = pd.to_numeric(temp[value_col], errors="coerce")
        out = temp.groupby("__period__")[value_col].agg(aggregation).reset_index()
        value_name = str(value_col)
    return out.sort_values("__period__"), value_name, label


def infer_chart_request(df, request):
    """Convert natural-language visualization requests into an explicit chart specification."""
    if df is None or df.empty:
        return {"type": "bar", "category": None, "value": None, "date": None, "top_n": None}

    raw = str(request or "")
    text = normalize_for_match(raw)
    numeric_cols = find_numeric_columns(df)
    date_cols = find_date_columns(df)
    dimension_candidates = _viz_dimension_candidates(df)
    chart_type = _viz_chart_type(raw, df)
    top_n = _extract_top_n_viz(raw)
    aggregation = _viz_aggregation(raw)
    horizontal = any(x in text for x in ["horizontal bar", "horizontal bars", "horizontal column"])
    stacked = "stacked" in text
    percent_stacked = any(x in text for x in ["100% stacked", "percentage stacked", "percent stacked"])

    dimensions = _extract_by_dimensions_viz(raw)
    category_col = None
    second_category_col = None

    if dimensions:
        category_col = _resolve_viz_column(df, dimensions[0], dimension_candidates)
        if len(dimensions) > 1:
            second_category_col = _resolve_viz_column(df, dimensions[1], dimension_candidates)

    metric_phrase = _viz_explicit_metric_phrase(raw)
    value_col = _resolve_viz_measure(df, raw, metric_phrase)

    # Handle "revenue by city" and similar phrases even when the parser sees
    # both sides as possible dimensions.
    by_match = re.search(r"\bby\b", raw, re.I)
    if by_match:
        left = _viz_clean_phrase(raw[:by_match.start()])
        right = _viz_clean_phrase(raw[by_match.end():])
        left_num = _resolve_viz_column(df, left, numeric_cols)
        right_num = _resolve_viz_column(df, right, numeric_cols)
        left_dim = _resolve_viz_column(df, left, dimension_candidates)
        right_dim = _resolve_viz_column(df, right, dimension_candidates)
        if right_dim and left_num:
            category_col, value_col = right_dim, left_num
        elif left_dim and right_num:
            category_col, value_col = left_dim, right_num
        elif right_dim:
            category_col = right_dim

    if chart_type == "scatter":
        left, right = _extract_vs_viz(raw)
        x_col = _resolve_viz_column(df, left, numeric_cols) if left else None
        y_col = _resolve_viz_column(df, right, numeric_cols) if right else None
        if not x_col and numeric_cols:
            x_col = numeric_cols[0]
        if not y_col:
            y_col = next((c for c in numeric_cols if c != x_col), None)
    else:
        x_col = y_col = None

    # Explicit "orders/count/number of" means row counts, not SUM(OrderID).
    if aggregation == "count":
        value_col = None

    # Date requests should use the date dimension when the user says trend/over time.
    date_col = _resolve_viz_column(df, _viz_clean_phrase(raw), date_cols)
    if not date_col and date_cols:
        date_col = date_cols[0]
    if chart_type in {"time", "area"} and date_cols:
        date_col = date_col or date_cols[0]

    if chart_type == "heatmap":
        return {
            "type": "heatmap", "category": category_col, "value": value_col,
            "date": date_col, "top_n": top_n, "aggregation": aggregation,
            "horizontal": horizontal, "stacked": stacked, "percent_stacked": percent_stacked,
            "second_category": second_category_col, "x": x_col, "y": y_col,
        }

    return {
        "type": chart_type, "category": category_col, "value": value_col,
        "date": date_col, "top_n": top_n, "aggregation": aggregation,
        "horizontal": horizontal, "stacked": stacked, "percent_stacked": percent_stacked,
        "second_category": second_category_col, "x": x_col, "y": y_col,
    }


def add_visualization(fig, chart_data, chart_name, request, chart_type):
    """Persist a visualization as an independent, deletable object."""
    visualization_id = uuid4().hex[:10]
    image_bytes = figure_to_bytes(fig)
    item = {
        "id": visualization_id,
        "type": "chart",
        "chart_type": chart_type,
        "name": chart_name,
        "request": request,
        "image": image_bytes,
        "data": chart_data.copy() if isinstance(chart_data, pd.DataFrame) else chart_data,
    }
    st.session_state.visualizations.append(item)
    st.session_state.report_items = [item.copy() for item in st.session_state.visualizations]
    st.session_state.last_chart = image_bytes
    st.session_state.last_chart_name = chart_name
    st.session_state.last_chart_data = chart_data
    return item


def delete_visualization(visualization_id):
    before = len(st.session_state.visualizations)
    st.session_state.visualizations = [
        item for item in st.session_state.visualizations if item.get("id") != visualization_id
    ]
    st.session_state.report_items = [item.copy() for item in st.session_state.visualizations]
    return len(st.session_state.visualizations) < before


def create_visualization_from_request(df, request):
    """Create a chart from natural language with explicit data/axis resolution."""
    if df is None or df.empty:
        return None, None, None, "No data available for visualization."

    spec = infer_chart_request(df, request)
    chart_type = spec["type"]
    category_col = spec.get("category")
    value_col = spec.get("value")
    date_col = spec.get("date")
    top_n = spec.get("top_n")
    aggregation = spec.get("aggregation", "sum")

    try:
        # -------------------- SCATTER --------------------
        if chart_type == "scatter":
            numeric_cols = find_numeric_columns(df)
            x_col = spec.get("x") or (numeric_cols[0] if numeric_cols else None)
            y_col = spec.get("y") or next((c for c in numeric_cols if c != x_col), None)
            if not x_col or not y_col:
                return None, None, None, "A scatter chart needs at least two numeric columns."
            chart_data = df[[x_col, y_col]].apply(pd.to_numeric, errors="coerce").dropna()
            fig, ax = plt.subplots(figsize=(10, 5))
            ax.scatter(chart_data[x_col], chart_data[y_col], alpha=0.75)
            ax.set_xlabel(str(x_col)); ax.set_ylabel(str(y_col)); ax.set_title(f"{y_col} vs {x_col}")
            return fig, chart_data, f"{y_col}_vs_{x_col}", None

        # -------------------- HEATMAP --------------------
        if chart_type == "heatmap":
            numeric = find_numeric_columns(df)
            if len(numeric) < 2:
                return None, None, None, "A correlation heatmap needs at least two numeric columns."
            corr = df[numeric].corr(numeric_only=True)
            fig, ax = plt.subplots(figsize=(10, 7))
            image = ax.imshow(corr.values, aspect="auto")
            ax.set_xticks(range(len(corr.columns))); ax.set_xticklabels(corr.columns, rotation=45, ha="right")
            ax.set_yticks(range(len(corr.index))); ax.set_yticklabels(corr.index)
            for i in range(len(corr.index)):
                for j in range(len(corr.columns)):
                    val = corr.iloc[i, j]
                    if pd.notna(val):
                        ax.text(j, i, f"{val:.2f}", ha="center", va="center")
            ax.set_title("Correlation Heatmap")
            fig.colorbar(image, ax=ax, fraction=0.046, pad=0.04)
            plt.tight_layout()
            return fig, corr.reset_index(), "Correlation_Heatmap", None

        # -------------------- HISTOGRAM --------------------
        if chart_type == "histogram":
            numeric = find_numeric_columns(df)
            value_col = value_col or (numeric[0] if numeric else None)
            if not value_col:
                return None, None, None, "A histogram needs a numeric column."
            chart_data = pd.to_numeric(df[value_col], errors="coerce").dropna().to_frame(name=value_col)
            fig, ax = plt.subplots(figsize=(10, 5))
            ax.hist(chart_data[value_col], bins="auto")
            ax.set_xlabel(str(value_col)); ax.set_ylabel("Frequency"); ax.set_title(f"Distribution of {value_col}")
            return fig, chart_data, f"{value_col}_Distribution", None

        # -------------------- BOX / VIOLIN --------------------
        if chart_type in {"box", "violin"}:
            numeric = find_numeric_columns(df)
            if not numeric:
                return None, None, None, "A box/violin plot needs a numeric column."
            if category_col:
                temp = df[[category_col, value_col or numeric[0]]].copy()
                val = value_col or numeric[0]
                temp[val] = pd.to_numeric(temp[val], errors="coerce")
                temp = temp.dropna(subset=[val])
                groups = [g[val].values for _, g in temp.groupby(category_col)]
                labels = [str(k) for k, _ in temp.groupby(category_col)]
                if top_n:
                    pairs = sorted(zip(labels, groups), key=lambda x: len(x[1]), reverse=True)[:top_n]
                    labels, groups = zip(*pairs) if pairs else ([], [])
                chart_data = temp
            else:
                val = value_col or numeric[0]
                groups = [pd.to_numeric(df[val], errors="coerce").dropna().values]
                labels = [str(val)]
                chart_data = df[[val]].copy()
            fig, ax = plt.subplots(figsize=(10, 5))
            if chart_type == "box":
                ax.boxplot(groups, labels=labels)
            else:
                ax.violinplot(groups, showmeans=True)
                ax.set_xticks(range(1, len(labels) + 1)); ax.set_xticklabels(labels, rotation=45, ha="right")
            ax.set_ylabel(str(val)); ax.set_title(f"{chart_type.title()} Plot of {val}" + (f" by {category_col}" if category_col else ""))
            plt.tight_layout()
            return fig, chart_data, f"{chart_type.title()}_{val}", None

        # -------------------- TIME / AREA --------------------
        if chart_type in {"time", "area"}:
            if not date_col:
                return None, None, None, "A time/area chart needs a date or time column."
            if aggregation != "count" and not value_col:
                return None, None, None, "A time/area chart needs a numeric measure or a count request."
            chart_data, value_name, period_label = _viz_prepare_time(df, date_col, value_col, aggregation, request)
            if chart_data.empty:
                return None, None, None, "No valid date values were found for the chart."
            fig, ax = plt.subplots(figsize=(11, 5))
            x = chart_data.iloc[:, 0]
            y = chart_data.iloc[:, 1]
            if chart_type == "area":
                ax.fill_between(x, y, alpha=0.35)
                ax.plot(x, y)
            else:
                ax.plot(x, y, marker="o")
            ax.set_xlabel(period_label); ax.set_ylabel(value_name); ax.set_title(f"{period_label} {value_name} Trend")
            plt.xticks(rotation=45, ha="right")
            return fig, chart_data, f"{period_label}_{value_name}_Trend", None

        # -------------------- PIE / DONUT --------------------
        if chart_type in {"pie", "donut"}:
            if not category_col:
                return None, None, None, "A pie/donut chart needs a category or identifier column."
            chart_data, value_name = _viz_build_grouped(df, category_col, value_col, aggregation, top_n)
            if chart_data.empty:
                return None, None, None, "There is no data available for this pie/donut chart."
            # A pie chart becomes unreadable with many slices. If no top-N was
            # requested, keep the largest slices and combine the remainder.
            if top_n is None and len(chart_data) > 12:
                keep = chart_data.head(10).copy()
                other = chart_data.iloc[10:][value_name].sum()
                if other != 0:
                    keep = pd.concat([keep, pd.DataFrame({category_col: ["Other"], value_name: [other]})], ignore_index=True)
                chart_data = keep
            fig, ax = plt.subplots(figsize=(9, 7))
            kwargs = {"labels": chart_data[category_col].astype(str), "autopct": "%1.1f%%", "startangle": 90}
            if chart_type == "donut":
                ax.pie(chart_data[value_name], wedgeprops={"width": 0.42}, **kwargs)
            else:
                ax.pie(chart_data[value_name], **kwargs)
            ax.set_title(f"{value_name} by {category_col}")
            plt.tight_layout()
            return fig, chart_data, f"{value_name}_by_{category_col}_{chart_type.title()}", None

        # -------------------- WATERFALL --------------------
        if chart_type == "waterfall":
            if not category_col:
                return None, None, None, "A waterfall chart needs a category column."
            chart_data, value_name = _viz_build_grouped(df, category_col, value_col, aggregation, top_n)
            if chart_data.empty:
                return None, None, None, "No data is available for the waterfall chart."
            vals = chart_data[value_name].astype(float).tolist()
            labels = chart_data[category_col].astype(str).tolist()
            cumulative = np.cumsum([0] + vals[:-1])
            fig, ax = plt.subplots(figsize=(11, 5))
            for i, (start, val) in enumerate(zip(cumulative, vals)):
                ax.bar(i, val, bottom=start if val >= 0 else start + val)
            ax.plot(range(len(vals)), np.cumsum(vals), marker="o")
            ax.set_xticks(range(len(labels))); ax.set_xticklabels(labels, rotation=45, ha="right")
            ax.set_ylabel(value_name); ax.set_title(f"Waterfall of {value_name} by {category_col}")
            plt.tight_layout()
            return fig, chart_data, f"{value_name}_Waterfall", None

        # -------------------- STACKED BAR --------------------
        if chart_type == "stacked_bar" or spec.get("stacked"):
            second = spec.get("second_category")
            if not category_col or not second:
                # If no second dimension was named, use a likely status/segment column.
                second = next((c for c in _viz_dimension_candidates(df) if c != category_col and any(k in normalize_for_match(c) for k in ["status", "type", "category", "segment", "payment", "gender"])), None)
            if not category_col or not second:
                return None, None, None, "A stacked bar chart needs two dimensions, for example 'revenue by city and hotel'."
            if aggregation == "count":
                grouped = df.groupby([category_col, second], dropna=False).size().reset_index(name="Count")
                value_name = "Count"
            else:
                if not value_col:
                    return None, None, None, "A stacked bar chart needs a numeric measure."
                temp = df[[category_col, second, value_col]].copy()
                temp[value_col] = pd.to_numeric(temp[value_col], errors="coerce")
                temp = temp.dropna(subset=[value_col])
                grouped = temp.groupby([category_col, second], dropna=False)[value_col].agg(aggregation).reset_index()
                value_name = str(value_col)
            pivot = grouped.pivot(index=category_col, columns=second, values=value_name).fillna(0)
            if top_n:
                pivot = pivot.loc[pivot.sum(axis=1).sort_values(ascending=False).head(top_n).index]
            if spec.get("percent_stacked"):
                pivot = pivot.div(pivot.sum(axis=1).replace(0, np.nan), axis=0) * 100
                value_name = "%"
            chart_data = pivot.reset_index()
            fig, ax = plt.subplots(figsize=(11, 6))
            pivot.plot(kind="bar", stacked=True, ax=ax)
            ax.set_xlabel(str(category_col)); ax.set_ylabel(value_name); ax.set_title(f"{value_name} by {category_col}, split by {second}")
            plt.xticks(rotation=45, ha="right")
            return fig, chart_data, f"{value_name}_by_{category_col}_Stacked", None

        # -------------------- BAR / COLUMN --------------------
        if chart_type == "bar":
            if not category_col:
                # If user only asked for a chart, automatically choose a useful dimension.
                dims = _viz_dimension_candidates(df)
                category_col = dims[0] if dims else None
            if not category_col:
                return None, None, None, "A bar chart needs a category, text, date, or identifier column."
            chart_data, value_name = _viz_build_grouped(df, category_col, value_col, aggregation, top_n)
            if chart_data.empty:
                return None, None, None, "No data is available for this bar chart."
            fig, ax = plt.subplots(figsize=(11, 6))
            labels = chart_data[category_col].astype(str)
            values = chart_data[value_name]
            if spec.get("horizontal"):
                ax.barh(labels, values)
                ax.invert_yaxis()
                ax.set_xlabel(value_name); ax.set_ylabel(str(category_col))
            else:
                ax.bar(labels, values)
                ax.set_xlabel(str(category_col)); ax.set_ylabel(value_name)
                plt.xticks(rotation=45, ha="right")
            prefix = f"Top {top_n} " if top_n else ""
            direction = "Horizontal " if spec.get("horizontal") else ""
            ax.set_title(f"{direction}{prefix}{value_name} by {category_col}")
            plt.tight_layout()
            return fig, chart_data, f"{value_name}_by_{category_col}" + (f"_Top{top_n}" if top_n else ""), None

        # -------------------- DEFAULT COUNT --------------------
        dims = _viz_dimension_candidates(df)
        category_col = category_col or (dims[0] if dims else None)
        if not category_col:
            return None, None, None, "No categorical or identifier column is available for visualization."
        chart_data, value_name = _viz_build_grouped(df, category_col, None, "count", top_n)
        fig, ax = plt.subplots(figsize=(11, 6))
        ax.bar(chart_data[category_col].astype(str), chart_data[value_name])
        ax.set_xlabel(str(category_col)); ax.set_ylabel("Count"); ax.set_title(f"Count by {category_col}")
        plt.xticks(rotation=45, ha="right")
        plt.tight_layout()
        return fig, chart_data, f"{category_col}_Count", None

    except Exception as exc:
        return None, None, None, f"Could not create visualization: {exc}"


def render_visualization_manager():
    """Render every saved visualization with an independent Delete button."""
    visualizations = st.session_state.get("visualizations", [])
    st.divider()
    st.subheader("📊 Visualizations")
    if not visualizations:
        st.info("No visualizations yet. Try: Show a horizontal bar chart of revenue by hotel")
        return
    st.caption(f"{len(visualizations)} visualization(s) saved. Each visualization can be deleted independently.")
    for index, item in enumerate(visualizations, start=1):
        viz_id = item.get("id")
        title = item.get("name", f"Visualization {index}")
        request = item.get("request", "")
        with st.container(border=True):
            header_col, delete_col = st.columns([8, 1])
            with header_col:
                st.markdown(f"### {index}. {title}")
                if request:
                    st.caption(f"Request: {request}")
            with delete_col:
                if st.button("🗑️", key=f"delete_viz_{viz_id}", help="Delete this visualization"):
                    delete_visualization(viz_id)
                    st.rerun()
            image_bytes = item.get("image")
            if image_bytes:
                st.image(image_bytes, use_container_width=True)
            chart_data = item.get("data")
            if isinstance(chart_data, pd.DataFrame) and not chart_data.empty:
                with st.expander("View chart data"):
                    st.dataframe(chart_data, use_container_width=True, hide_index=True)


def create_chart(df, request):
    """Backward-compatible wrapper around the visualization engine."""
    fig, chart_data, chart_name, error = create_visualization_from_request(df, request)
    if error:
        return None, None, error
    return fig, chart_data, chart_name


def figure_to_bytes( 
    fig 
): 
 
    buffer = io.BytesIO() 
 
    fig.savefig( 
        buffer, 
        format="png", 
        dpi=150, 
        bbox_inches="tight" 
    ) 
 
    buffer.seek(0) 
 
    return buffer.getvalue() 
 
 
# ============================================================ 
# BUSINESS INSIGHTS 
# ============================================================ 
 
def generate_business_insights( 
    df 
): 
 
    insights = [] 
 
    if df is None or df.empty: 
 
        return insights 
 
    # -------------------------------------------------------- 
    # ROW COUNT 
    # -------------------------------------------------------- 
 
    insights.append( 
        f"The dataset contains " 
        f"{len(df):,} rows and " 
        f"{len(df.columns):,} columns." 
    ) 
 
    # -------------------------------------------------------- 
    # MISSING 
    # -------------------------------------------------------- 
 
    missing = int( 
        df.isna() 
        .sum() 
        .sum() 
    ) 
 
    if missing == 0: 
 
        insights.append( 
            "No missing values remain " 
            "in the current dataset." 
        ) 
 
    else: 
 
        insights.append( 
            f"The current dataset contains " 
            f"{missing:,} missing cell(s)." 
        ) 
 
    # -------------------------------------------------------- 
    # DUPLICATES 
    # -------------------------------------------------------- 
 
    duplicates = int( 
        df.duplicated().sum() 
    ) 
 
    if duplicates == 0: 
 
        insights.append( 
            "No duplicate rows " 
            "are present." 
        ) 
 
    else: 
 
        insights.append( 
            f"There are " 
            f"{duplicates:,} duplicate rows." 
        ) 
 
    # -------------------------------------------------------- 
    # NUMERIC 
    # -------------------------------------------------------- 
 
    numeric_cols = ( 
        find_numeric_columns( 
            df 
        ) 
    ) 
 
    for col in numeric_cols[:5]: 
 
        total = df[col].sum() 
 
        average = df[col].mean() 
 
        insights.append( 
            f"{col}: total = " 
            f"{total:,.2f}, " 
            f"average = " 
            f"{average:,.2f}." 
        ) 
 
    # -------------------------------------------------------- 
    # CATEGORICAL 
    # -------------------------------------------------------- 
 
    categorical_cols = [ 
        col 
        for col in df.columns 
        if ( 
            df[col].dtype == "object" 
            or 
            pd.api.types.is_string_dtype( 
                df[col] 
            ) 
        ) 
    ] 
 
    for col in categorical_cols[:3]: 
 
        values = ( 
            df[col] 
            .dropna() 
            .value_counts() 
        ) 
 
        if values.empty: 
            continue 
 
        top_value = ( 
            values.index[0] 
        ) 
 
        top_count = ( 
            values.iloc[0] 
        ) 
 
        insights.append( 
            f"Most frequent value in " 
            f"'{col}' is '{top_value}' " 
            f"with {top_count:,} record(s)." 
        ) 
 
    return insights 
 
 
# ============================================================ 
# EXCEL REPORT GENERATION 
# ============================================================ 
 
def generate_excel_report(): 
 
    output = io.BytesIO() 
 
    with pd.ExcelWriter( 
        output, 
        engine="openpyxl" 
    ) as writer: 
 
        # ---------------------------------------------------- 
        # RAW DATA 
        # ---------------------------------------------------- 
 
        if ( 
            st.session_state.raw_df 
            is not None 
        ): 
 
            st.session_state.raw_df.to_excel( 
                writer, 
                sheet_name="Raw Data", 
                index=False 
            ) 
 
        # ---------------------------------------------------- 
        # CLEANED DATA 
        # ---------------------------------------------------- 
 
        if ( 
            st.session_state.cleaned_df 
            is not None 
        ): 
 
            st.session_state.cleaned_df.to_excel( 
                writer, 
                sheet_name="Cleaned Data", 
                index=False 
            ) 
 
        # ---------------------------------------------------- 
        # TRANSFORMED DATA 
        # ---------------------------------------------------- 
 
        if ( 
            st.session_state.transformed_df 
            is not None 
        ): 
 
            st.session_state.transformed_df.to_excel( 
                writer, 
                sheet_name="Transformed Data", 
                index=False 
            ) 
 
        # ---------------------------------------------------- 
        # PIVOT 
        # ---------------------------------------------------- 
 
        if ( 
            st.session_state.pivot_df 
            is not None 
            and 
            not st.session_state.pivot_df.empty 
        ): 
 
            st.session_state.pivot_df.to_excel( 
                writer, 
                sheet_name="Pivot Table", 
                index=False 
            ) 
 
        # ---------------------------------------------------- 
        # VALIDATION 
        # ---------------------------------------------------- 
 
        if ( 
            st.session_state.validation_df 
            is not None 
            and 
            not st.session_state.validation_df.empty 
        ): 
 
            st.session_state.validation_df.to_excel( 
                writer, 
                sheet_name="Validation", 
                index=False 
            ) 
 
        # ---------------------------------------------------- 
        # CLEANING LOG 
        # ---------------------------------------------------- 
 
        pd.DataFrame({ 
            "Cleaning Step": 
                st.session_state.cleaning_log 
        }).to_excel( 
            writer, 
            sheet_name="Cleaning Log", 
            index=False 
        ) 
 
        # ---------------------------------------------------- 
        # TRANSFORMATION LOG 
        # ---------------------------------------------------- 
 
        pd.DataFrame({ 
            "Transformation Step": 
                st.session_state.transformation_log 
        }).to_excel( 
            writer, 
            sheet_name="Transformation Log", 
            index=False 
        ) 
 
        # ---------------------------------------------------- 
        # ANALYSIS LOG 
        # ---------------------------------------------------- 
 
        pd.DataFrame({ 
            "Analysis": 
                st.session_state.analysis_log 
        }).to_excel( 
            writer, 
            sheet_name="Analysis Log", 
            index=False 
        ) 
 
        # ---------------------------------------------------- 
        # CALCULATED COLUMNS 
        # ---------------------------------------------------- 
 
        calculated_columns = ( 
            get_calculated_columns( 
                st.session_state.transformed_df 
            ) 
            if ( 
                st.session_state.transformed_df 
                is not None 
            ) 
            else [] 
        ) 
 
        pd.DataFrame({ 
            "Calculated Columns": 
                calculated_columns 
        }).to_excel( 
            writer, 
            sheet_name="Calculated Columns", 
            index=False 
        ) 
 
        # ---------------------------------------------------- 
        # ANALYSIS RESULTS 
        # ---------------------------------------------------- 
 
        analysis_tables = [] 
 
        for item in ( 
            st.session_state.analysis_results 
        ): 
 
            if isinstance( 
                item, 
                pd.DataFrame 
            ): 
 
                analysis_tables.append( 
                    item 
                ) 
 
        if analysis_tables: 
 
            start_row = 0 
 
            for table in analysis_tables: 
 
                table.to_excel( 
                    writer, 
                    sheet_name="Analysis Results", 
                    startrow=start_row, 
                    index=False 
                ) 
 
                start_row += ( 
                    len(table) + 3 
                ) 
 
        else: 
 
            pd.DataFrame({ 
                "Message": [ 
                    "No analysis tables generated." 
                ] 
            }).to_excel( 
                writer, 
                sheet_name="Analysis Results", 
                index=False 
            ) 
 
        # ---------------------------------------------------- 
        # BUSINESS INSIGHTS 
        # ---------------------------------------------------- 
 
        insights = ( 
            generate_business_insights( 
                get_working_df() 
            ) 
        ) 
 
        pd.DataFrame({ 
            "Business Insight": 
                insights 
        }).to_excel( 
            writer, 
            sheet_name="Business Insights", 
            index=False 
        ) 
 
        # ---------------------------------------------------- 
        # REPORT INFO 
        # ---------------------------------------------------- 
 
        pd.DataFrame({ 
            "Report Information": [ 
                "AI Excel Analyst Report", 
 
                f"Generated: " 
                f"{datetime.now()}", 
 
                ( 
                    "Source file: " 
                    f"{st.session_state.uploaded_filename}" 
                ) 
            ] 
        }).to_excel( 
            writer, 
            sheet_name="Report Info", 
            index=False 
        ) 
 
    # ======================================================== 
    # ADD VISUALIZATIONS 
    # ======================================================== 
 
    output.seek(0) 
 
    workbook = load_workbook( 
        output 
    ) 
 
    if "Visualizations" in ( 
        workbook.sheetnames 
    ): 
 
        del workbook[ 
            "Visualizations" 
        ] 
 
    visualization_sheet = ( 
        workbook.create_sheet( 
            "Visualizations" 
        ) 
    ) 
 
    visualization_sheet["A1"] = ( 
        "AI Excel Analyst - Visualizations" 
    ) 
 
    visualization_sheet["A2"] = ( 
        "Charts generated during analysis" 
    ) 
 
    row_position = 4 
 
    for item in ( 
        st.session_state.get( 
            "visualizations", 
            st.session_state.get("report_items", []) 
        ) 
    ): 
 
        if not isinstance( 
            item, 
            dict 
        ): 
            continue 
 
        if item.get( 
            "type" 
        ) != "chart": 
            continue 
 
        image_bytes = ( 
            item.get( 
                "image" 
            ) 
        ) 
 
        if not image_bytes: 
            continue 
 
        image_stream = io.BytesIO( 
            image_bytes 
        ) 
 
        image = XLImage( 
            image_stream 
        ) 
 
        visualization_sheet[ 
            f"A{row_position}" 
        ] = item.get( 
            "name", 
            "Chart" 
        ) 
 
        image.anchor = ( 
            f"A{row_position + 1}" 
        ) 
 
        visualization_sheet.add_image( 
            image 
        ) 
 
        row_position += 25 
 
        chart_data = item.get( 
            "data" 
        ) 
 
        if ( 
            isinstance( 
                chart_data, 
                pd.DataFrame 
            ) 
            and 
            not chart_data.empty 
        ): 
 
            data_start = ( 
                row_position 
            ) 
 
            # Header 
            for c_index, column in enumerate( 
                chart_data.columns, 
                start=1 
            ): 
 
                visualization_sheet.cell( 
                    row=data_start, 
                    column=c_index 
                ).value = str( 
                    column 
                ) 
 
            # Data 
            for r_index, row_values in enumerate( 
                chart_data.itertuples( 
                    index=False, 
                    name=None 
                ), 
                start=data_start + 1 
            ): 
 
                for c_index, value in enumerate( 
                    row_values, 
                    start=1 
                ): 
 
                    if pd.isna( 
                        value 
                    ): 
 
                        value = None 
 
                    visualization_sheet.cell( 
                        row=r_index, 
                        column=c_index 
                    ).value = value 
 
            row_position = ( 
                data_start 
                + 
                len(chart_data) 
                + 
                4 
            ) 
 
    # ======================================================== 
    # SAVE 
    # ======================================================== 
 
    final_output = io.BytesIO() 
 
    workbook.save( 
        final_output 
    ) 
 
    final_output.seek(0) 
 
    return final_output.getvalue() 
 
 
# ============================================================ 
# SIDEBAR 
# ============================================================ 
 
with st.sidebar: 
 
    st.title( 
        "📊 AI Excel Analyst" 
    ) 
 
    st.caption( 
        "Analyze any Excel dataset" 
    ) 
 
    st.divider() 
 
    stages = [ 
        ("1", "Upload"), 
        ("2", "Inspect"), 
        ("3", "Clean"), 
        ("4", "Validate"), 
        ("5", "Transform"), 
        ("6", "Pivot"), 
        ("7", "Analyze"), 
        ("8", "Report") 
    ] 
 
    for number, name in stages: 
 
        st.write( 
            f"**{number}. {name}**" 
        ) 
 
 
# ============================================================ 
# HEADER 
# ============================================================ 
 
st.title( 
    "📊 AI Excel Analyst" 
) 
 
st.caption( 
    "Upload an Excel file and analyze it step by step." 
) 
 
 
# ============================================================ 
# STAGE 1 — UPLOAD 
# ============================================================ 
 
def answer_business_question(df, question):
    """Answer common business questions from the current DataFrame."""
    if df is None or df.empty:
        return "Please upload a dataset first."

    data = df.copy()
    data.columns = [re.sub(r"[^a-z0-9]+", "_", str(c).lower()).strip("_") for c in data.columns]
    q = str(question).lower().strip()
    qc = re.sub(r"[^a-z0-9]", "", q)

    def col(*names):
        normalized = {re.sub(r"[^a-z0-9]+", "_", n.lower()).strip("_") for n in names}
        compact = {re.sub(r"[^a-z0-9]", "", n.lower()) for n in names}
        for c in data.columns:
            if c in normalized or re.sub(r"[^a-z0-9]", "", c) in compact:
                return c
        return None

    hotel = col("hotel", "hotel_name", "restaurant", "restaurant_name")
    food = col("food_item", "food", "fooditem", "item", "item_name", "product", "dish", "menu_item")
    customer = col("customer_id", "customer", "customer_name", "customerid")
    payment = col("payment_method", "payment", "payment_type", "paymentmethod")
    partner = col("delivery_partner", "delivery_partner_id", "partner_id", "partner", "delivery_person", "deliveryperson")
    price = col("price", "unit_price", "selling_price")
    qty = col("qty", "quantity", "units")
    status = col("status", "order_status", "delivery_status", "orderstatus")
    date = col("order_date", "date", "order_datetime", "order_date_time", "orderdate", "datetime")
    revenue_col = col("revenue", "sales", "total_revenue", "total_sales")

    if revenue_col:
        data["_revenue"] = pd.to_numeric(data[revenue_col], errors="coerce").fillna(0)
    elif price and qty:
        data["_revenue"] = pd.to_numeric(data[price], errors="coerce").fillna(0) * pd.to_numeric(data[qty], errors="coerce").fillna(0)
    elif price:
        data["_revenue"] = pd.to_numeric(data[price], errors="coerce").fillna(0)
    else:
        data["_revenue"] = 0.0

    if status:
        sv = data[status].astype(str).str.strip().str.lower()
        data["_cancelled"] = sv.isin(["cancelled", "canceled", "cancel", "failed"])
    else:
        data["_cancelled"] = False
    data["_delivered"] = ~data["_cancelled"]

    def money(x):
        return f"{float(x):,.2f}"

    def rate(n, d):
        return f"{100*n/d:.1f}%" if d else "0.0%"

    def totals(by):
        return data.groupby(by, dropna=False)["_revenue"].sum().sort_values(ascending=False)

    def cancel_rates(by):
        return data.groupby(by, dropna=False)["_cancelled"].mean()

    # Open-ended questions must be handled before generic KPI matching.
    insight = any(x in q for x in [
        "business insight", "business insights", "give me insights", "give insights",
        "why is the cancellation", "why are the cancellations", "why is cancellation",
        "what should the business do", "how can the business", "improve revenue", "improve sales",
        "which hotel needs attention", "hotel needs attention"
    ])
    if insight:
        total = data["_revenue"].sum()
        cancelled_n = int(data["_cancelled"].sum())
        cancelled_rev = data.loc[data["_cancelled"], "_revenue"].sum()
        lines = [
            f"Total revenue is {money(total)} across {len(data)} orders.",
            f"{cancelled_n} of {len(data)} orders were cancelled ({rate(cancelled_n, len(data))}), with {money(cancelled_rev)} in cancelled-order revenue."
        ]
        if payment:
            pr = cancel_rates(payment).sort_values(ascending=False)
            if len(pr):
                p = pr.index[0]
                pc = int(data.loc[data[payment] == p, "_cancelled"].sum())
                pn = int((data[payment] == p).sum())
                lines.append(f"{p} has the highest payment-method cancellation rate: {pc} of {pn} orders ({rate(pc, pn)}).")
        if partner:
            rr = cancel_rates(partner).sort_values(ascending=False)
            if len(rr):
                p = rr.index[0]
                pc = int(data.loc[data[partner] == p, "_cancelled"].sum())
                pn = int((data[partner] == p).sum())
                lines.append(f"Delivery partner {p} has the highest cancellation rate: {pc} of {pn} orders ({rate(pc, pn)}).")
        if hotel:
            hr = cancel_rates(hotel).sort_values(ascending=False)
            if len(hr):
                lines.append(f"{hr.index[0]} has the highest hotel cancellation rate at {hr.iloc[0]*100:.1f}%.")
        if "improve revenue" in q or "improve sales" in q or "what should the business do" in q:
            actions = []
            if payment and len(cancel_rates(payment)):
                actions.append(f"Investigate the payment channel with the highest cancellation rate ({cancel_rates(payment).idxmax()}).")
            if partner and len(cancel_rates(partner)):
                actions.append(f"Review orders handled by delivery partner {cancel_rates(partner).idxmax()}.")
            if food and len(totals(food)):
                actions.append(f"Protect availability and sales of {totals(food).index[0]}, the highest-revenue food item.")
            actions.append("Reducing cancellations can recover revenue currently lost from cancelled orders.")
            lines.append("Potential actions based on the data: " + " ".join(actions))
        return "\n".join(f"• {x}" for x in lines)

    # Overall KPIs
    if "total revenue" in q or "overall revenue" in q or "total sales" in q:
        return f"Total revenue: {money(data['_revenue'].sum())}"
    if "average order value" in q or "aov" in q:
        return f"Average order value: {money(data['_revenue'].mean())}"
    if "fulfillment rate" in q or "fulfilment rate" in q:
        n = int(data["_delivered"].sum())
        return f"Fulfillment rate: {rate(n, len(data))} ({n} of {len(data)} orders)"
    if "cancellation rate" in q or "cancel rate" in q:
        n = int(data["_cancelled"].sum())
        return f"Cancellation rate: {rate(n, len(data))} ({n} of {len(data)} orders)"
    if "revenue lost" in q or "lost revenue" in q or "cancelled orders" in q and "revenue" in q:
        return f"Revenue lost due to cancelled orders: {money(data.loc[data['_cancelled'], '_revenue'].sum())}"
    if "revenue from delivered" in q or "delivered revenue" in q:
        return f"Revenue from delivered orders only: {money(data.loc[data['_delivered'], '_revenue'].sum())}"
    if "average quantity" in q or "average units per order" in q:
        if not qty:
            return "The dataset does not contain a quantity column."
        s = pd.to_numeric(data[qty], errors="coerce").fillna(0)
        return f"Average quantity per order: {s.mean():.2f} ({s.sum():g} items in total)"

    # Hotel questions
    if hotel and "hotel" in q:
        if "most orders" in q or "highest number of orders" in q:
            s = data[hotel].value_counts(dropna=False)
            return f"Hotel with most orders: {s.index[0]} ({int(s.iloc[0])} orders)" if len(s) else None
        if "cancellation" in q or "cancel" in q:
            s = cancel_rates(hotel).sort_values(ascending="lowest" not in q and "least" not in q)
            if len(s):
                return f"Hotel with {'lowest' if ('lowest' in q or 'least' in q) else 'highest'} cancellation rate: {s.index[0]} ({s.iloc[0]*100:.1f}%)"
        if "average order value" in q or "aov" in q:
            s = data.groupby(hotel)["_revenue"].mean().sort_values(ascending=False)
            return f"Hotel with highest average order value: {s.index[0]} ({money(s.iloc[0])})" if len(s) else None
        if "revenue" in q or "sales" in q:
            s = totals(hotel)
            return f"Hotel with highest revenue: {s.index[0]} ({money(s.iloc[0])})" if len(s) else None

    # Food questions
    if food and any(x in q for x in ["food", "item", "dish", "menu", "product"]):
        if any(x in q for x in ["zero cancellation", "zero cancellations", "no cancellation", "no cancellations"]):
            s = cancel_rates(food)
            z = s[s == 0]
            return "Food items with zero cancellations: " + ", ".join(map(str, z.index)) if len(z) else "No food items have zero cancellations."
        if "cancellation" in q or "cancel" in q:
            s = cancel_rates(food).sort_values(ascending=False)
            return f"Food item with highest cancellation rate: {s.index[0]} ({s.iloc[0]*100:.1f}%)" if len(s) else None
        if "ordered the most" in q or "most ordered" in q or "most sold" in q or "most units" in q:
            if not qty:
                return "The dataset does not contain a quantity column."
            s = data.assign(_qty=pd.to_numeric(data[qty], errors="coerce").fillna(0)).groupby(food)["_qty"].sum().sort_values(ascending=False)
            return f"Food item ordered the most: {s.index[0]} ({s.iloc[0]:g} units)" if len(s) else None
        if "average order value" in q or "highest average" in q or "aov" in q:
            s = data.groupby(food)["_revenue"].mean().sort_values(ascending=False)
            return f"Food item with highest average order value: {s.index[0]} ({money(s.iloc[0])})" if len(s) else None
        if "revenue" in q or "sales" in q:
            s = totals(food)
            return f"Food item with most revenue: {s.index[0]} ({money(s.iloc[0])})" if len(s) else None

    # Customer questions
    if customer and "customer" in q:
        counts = data[customer].value_counts(dropna=False)
        if "unique" in q:
            return f"Unique customers: {data[customer].nunique()}"
        if "more than one" in q or "multiple orders" in q or "repeat" in q:
            return f"Customers with more than one order: {int((counts > 1).sum())}"
        if "most orders" in q or "highest orders" in q:
            return f"Customer with most orders: {counts.index[0]} ({int(counts.iloc[0])} orders)" if len(counts) else None
        if "top 5" in q or "top five" in q:
            s = data.groupby(customer)["_revenue"].sum().sort_values(ascending=False).head(5)
            return "Top 5 customers by revenue:\n" + "\n".join(f"{k}: {money(v)}" for k, v in s.items()) if len(s) else None

    # Payment questions
    if payment and ("payment" in q or any(x in q for x in ["upi", "cash", "card"])):
        if "ever" in q and "cancel" in q:
            method = next((m for m in data[payment].dropna().unique() if str(m).lower() in q), None)
            if method is not None:
                n = int(data.loc[data[payment] == method, "_cancelled"].sum())
                return f"No. {method} orders have 0 cancellations." if n == 0 else f"Yes. {method} orders have {n} cancellation(s)."
        if "no cancellation" in q or "no cancellations" in q:
            z = cancel_rates(payment)
            z = z[z == 0]
            return "Payment methods with no cancellations: " + ", ".join(map(str, z.index)) if len(z) else "No payment methods have zero cancellations."
        if (
            "most used" in q
            or "used the most" in q
            or "most commonly used" in q
            or "most common" in q
            or "most popular" in q
        ):
            s = data[payment].value_counts(dropna=False)
            return f"Most used payment method: {s.index[0]} ({int(s.iloc[0])} orders)" if len(s) else None
        if "cancellation rate" in q or "cancel rate" in q:
            s = cancel_rates(payment).sort_values(ascending=False)
            if len(s):
                p = s.index[0]; n = int(data.loc[data[payment] == p, "_cancelled"].sum()); d = int((data[payment] == p).sum())
                return f"Payment method with highest cancellation rate: {p} ({n} of {d} orders, {rate(n,d)})"
        if "revenue" in q or "sales" in q:
            s = totals(payment)
            return f"Payment method with most revenue: {s.index[0]} ({money(s.iloc[0])})" if len(s) else None

    # Delivery partner questions
    if partner and "partner" in q:
        if "no cancellation" in q or "no cancellations" in q:
            z = cancel_rates(partner); z = z[z == 0]
            return "Delivery partners with no cancellations: " + ", ".join(map(str, z.index)) if len(z) else "No delivery partners have zero cancellations."
        if "cancellation rate" in q or "cancel rate" in q:
            s = cancel_rates(partner).sort_values(ascending=False)
            if len(s):
                p=s.index[0]; n=int(data.loc[data[partner]==p,"_cancelled"].sum()); d=int((data[partner]==p).sum())
                return f"Delivery partner with highest cancellation rate: {p} ({n} of {d} orders, {rate(n,d)})"
        if "revenue" in q or "sales" in q:
            s=totals(partner)
            return f"Delivery partner with most revenue: {s.index[0]} ({money(s.iloc[0])})" if len(s) else None

    # Time trends
    if date and any(x in q for x in ["day", "date", "daily", "trend"]):
        dts = pd.to_datetime(data[date], errors="coerce")
        daily = data.assign(_date=dts.dt.strftime("%Y-%m-%d")).groupby("_date")["_revenue"].sum()
        if len(daily):
            low = "lowest" in q or "minimum" in q
            value = daily.min() if low else daily.max()
            days = daily[daily == value].index
            label = "Lowest" if low else "Highest"
            return f"{label} revenue day(s): " + ", ".join(map(str, days)) + f" ({money(value)} each)"

    return None


if st.session_state.stage == "upload": 
 
    st.subheader( 
        "Upload Excel File" 
    ) 
 
    uploaded_file = st.file_uploader( 
        "Upload your Excel file", 
        type=[ 
            "xlsx", 
            "xls" 
        ] 
    ) 
 
    if uploaded_file is not None: 
 
        try: 
 
            file_bytes = ( 
                uploaded_file.getvalue() 
            ) 
 
            df = pd.read_excel( 
                io.BytesIO( 
                    file_bytes 
                ) 
            ) 
 
            st.session_state.raw_df = ( 
                df.copy() 
            ) 
 
            st.session_state.cleaned_df = ( 
                None 
            ) 
 
            st.session_state.transformed_df = ( 
                None 
            ) 
 
            st.session_state.uploaded_filename = ( 
                uploaded_file.name 
            ) 
 
            st.session_state.original_columns = ( 
                list(df.columns) 
            ) 
 
            st.session_state.calculated_columns = ( 
                [] 
            ) 
 
            st.session_state.inspection = ( 
                inspect_dataframe( 
                    df 
                ) 
            ) 
 
            st.session_state.cleaning_log = [] 
 
            st.session_state.transformation_log = [] 
 
            st.session_state.analysis_log = [] 
 
            st.session_state.validation_df = None 
 
            st.session_state.pivot_df = None 
 
            st.session_state.analysis_results = [] 
 
            st.session_state.report_items = [] 
            st.session_state.visualizations = [] 
 
            st.session_state.chat_messages = [] 
 
            st.session_state.current_result = None 
 
            st.session_state.last_chart = None 
 
            st.session_state.last_chart_name = "" 
 
            st.session_state.last_chart_data = None 
 
            st.session_state.report_generated = False 
 
            st.session_state.stage = ( 
                "inspect" 
            ) 
 
            st.rerun() 
 
        except Exception as e: 
 
            st.error( 
                f"Unable to read Excel file: {e}" 
            ) 
 
 
# ============================================================ 
# STAGE 2 — INSPECTION 
# ============================================================ 
 
elif st.session_state.stage == "inspect": 
 
    df = ( 
        st.session_state.raw_df 
    ) 
 
    st.subheader( 
        "Dataset Inspection" 
    ) 
 
    if df is None: 
 
        st.error( 
            "No dataset available." 
        ) 
 
    else: 
 
        col1, col2, col3, col4 = ( 
            st.columns(4) 
        ) 
 
        with col1: 
 
            st.metric( 
                "Rows", 
                len(df) 
            ) 
 
        with col2: 
 
            st.metric( 
                "Columns", 
                len(df.columns) 
            ) 
 
        with col3: 
 
            st.metric( 
                "Duplicate Rows", 
                int( 
                    df.duplicated().sum() 
                ) 
            ) 
 
        with col4: 
 
            st.metric( 
                "Missing Cells", 
                int( 
                    df.isna() 
                    .sum() 
                    .sum() 
                ) 
            ) 
 
        st.subheader( 
            "Column Information" 
        ) 
 
        column_info = pd.DataFrame({ 
 
            "Column": df.columns, 
 
            "Data Type": [ 
                str(dtype) 
                for dtype 
                in df.dtypes 
            ], 
 
            "Missing": [ 
                int( 
                    df[col].isna().sum() 
                ) 
                for col 
                in df.columns 
            ], 
 
            "Unique Values": [ 
                df[col].nunique( 
                    dropna=True 
                ) 
                for col 
                in df.columns 
            ] 
        }) 
 
        st.dataframe( 
            column_info, 
            use_container_width=True, 
            hide_index=True 
        ) 
 
        st.subheader( 
            "Data Preview" 
        ) 
 
        st.dataframe( 
            df.head(20), 
            use_container_width=True, 
            hide_index=True 
        ) 
 
        if st.button( 
            "Continue to Cleaning", 
            type="primary" 
        ): 
 
            st.session_state.stage = ( 
                "clean_choice" 
            ) 
 
            st.rerun() 
 
 
# ============================================================ 
# STAGE 3 — CLEANING CHOICE 
# ============================================================ 
 
elif st.session_state.stage == "clean_choice": 
 
    st.subheader( 
        "Choose Cleaning Method" 
    ) 
 
    st.write( 
        "Choose how you want to clean your Excel data." 
    ) 
 
    col1, col2 = st.columns(2) 
 
    with col1: 
 
        st.markdown( 
            "### 🤖 Automatic Cleaning" 
        ) 
 
        st.write( 
            "Automatically handles common " 
            "data-quality issues." 
        ) 
 
        if st.button( 
            "Run Automatic Cleaning", 
            type="primary" 
        ): 
 
            cleaned_df, log = ( 
                automatic_cleaning( 
                    st.session_state.raw_df 
                ) 
            ) 
 
            st.session_state.cleaned_df = ( 
                cleaned_df 
            ) 
 
            st.session_state.cleaning_log.extend( 
                log 
            ) 
 
            # The cleaned source columns become 
            # the protected source columns. 
            st.session_state.original_columns = ( 
                list( 
                    cleaned_df.columns 
                ) 
            ) 
 
            st.session_state.calculated_columns = [] 
 
            st.session_state.stage = ( 
                "cleaning_validation" 
            ) 
 
            st.rerun() 
 
    with col2: 
 
        st.markdown( 
            "### 🛠️ Manual Cleaning" 
        ) 
 
        st.write( 
            "Choose individual cleaning operations." 
        ) 
 
        if st.button( 
            "Open Manual Cleaning" 
        ): 
 
            st.session_state.stage = ( 
                "manual_cleaning" 
            ) 
 
            st.rerun() 
 
 
# ============================================================ 
# STAGE 3B — MANUAL CLEANING 
# ============================================================ 
 
elif st.session_state.stage == "manual_cleaning": 
 
    st.subheader( 
        "Manual Cleaning" 
    ) 
 
    df = ( 
        st.session_state.cleaned_df 
        if st.session_state.cleaned_df 
        is not None 
        else st.session_state.raw_df 
    ) 
 
    operation = st.selectbox( 
        "Select cleaning operation", 
        [ 
            "TRIM", 
            "CLEAN", 
            "UPPER", 
            "LOWER", 
            "PROPER", 
            "REMOVE DUPLICATES", 
            "REMOVE BLANK ROWS", 
            "FILL MISSING VALUES", 
            "REPLACE VALUES", 
            "CONVERT TO NUMBER", 
            "CONVERT TO DATE" 
        ] 
    ) 
 
    column = None 
 
    if operation not in [ 
        "REMOVE DUPLICATES", 
        "REMOVE BLANK ROWS" 
    ]: 
 
        column = st.selectbox( 
            "Select column", 
            list(df.columns) 
        ) 
 
    value = None 
 
    replacement = None 
 
    if operation == "FILL MISSING VALUES": 
 
        value = st.text_input( 
            "Value to use for missing cells", 
            value="Unknown" 
        ) 
 
    elif operation == "REPLACE VALUES": 
 
        value = st.text_input( 
            "Value to replace" 
        ) 
 
        replacement = st.text_input( 
            "Replacement value" 
        ) 
 
    if st.button( 
        "Apply Cleaning", 
        type="primary" 
    ): 
 
        cleaned_df, log = ( 
            apply_manual_cleaning( 
                df, 
                operation, 
                column, 
                value, 
                replacement 
            ) 
        ) 
 
        st.session_state.cleaned_df = ( 
            cleaned_df 
        ) 
 
        st.session_state.cleaning_log.extend( 
            log 
        ) 
 
        st.success( 
            "Cleaning operation applied." 
        ) 
 
        st.dataframe( 
            cleaned_df.head(20), 
            use_container_width=True, 
            hide_index=True 
        ) 
 
    st.divider() 
 
    if st.button( 
        "Cleaning Complete → Validate" 
    ): 
 
        if ( 
            st.session_state.cleaned_df 
            is None 
        ): 
 
            st.session_state.cleaned_df = ( 
                st.session_state.raw_df.copy() 
            ) 
 
        st.session_state.original_columns = ( 
            list( 
                st.session_state.cleaned_df.columns 
            ) 
        ) 
 
        st.session_state.calculated_columns = [] 
 
        st.session_state.stage = ( 
            "cleaning_validation" 
        ) 
 
        st.rerun() 
 
 
# ============================================================ 
# STAGE 4 — CLEANING VALIDATION 
# ============================================================ 
 
elif st.session_state.stage == "cleaning_validation": 
 
    st.subheader( 
        "🔎 Cleaning Validation" 
    ) 
 
    raw_df = ( 
        st.session_state.raw_df 
    ) 
 
    cleaned_df = ( 
        st.session_state.cleaned_df 
    ) 
 
    if ( 
        raw_df is None 
        or 
        cleaned_df is None 
    ): 
 
        st.error( 
            "Raw or cleaned data is unavailable." 
        ) 
 
    else: 
 
        validation_df, overall_status = ( 
            validate_cleaning( 
                raw_df, 
                cleaned_df 
            ) 
        ) 
 
        st.session_state.validation_df = ( 
            validation_df 
        ) 
 
        passed = int( 
            ( 
                validation_df["Status"] 
                == 
                "PASS" 
            ).sum() 
        ) 
 
        warnings = int( 
            ( 
                validation_df["Status"] 
                == 
                "WARNING" 
            ).sum() 
        ) 
 
        information = int( 
            ( 
                validation_df["Status"] 
                == 
                "INFO" 
            ).sum() 
        ) 
 
        col1, col2, col3 = ( 
            st.columns(3) 
        ) 
 
        with col1: 
 
            st.metric( 
                "Passed", 
                passed 
            ) 
 
        with col2: 
 
            st.metric( 
                "Warnings", 
                warnings 
            ) 
 
        with col3: 
 
            st.metric( 
                "Information", 
                information 
            ) 
 
        st.dataframe( 
            validation_df, 
            use_container_width=True, 
            hide_index=True 
        ) 
 
        if overall_status == "PASS": 
 
            st.success( 
                "Cleaning validation completed successfully." 
            ) 
 
        else: 
 
            st.warning( 
                "Some validation checks require attention." 
            ) 
 
        with st.expander( 
            "View cleaned data" 
        ): 
 
            st.dataframe( 
                cleaned_df, 
                use_container_width=True, 
                hide_index=True 
            ) 
 
        col1, col2 = st.columns(2) 
 
        with col1: 
 
            if st.button( 
                "Run Validation Again" 
            ): 
 
                st.rerun() 
 
        with col2: 
 
            if st.button( 
                "Continue to Transform", 
                type="primary" 
            ): 
 
                st.session_state.transformed_df = ( 
                    cleaned_df.copy() 
                ) 
 
                # Protect source columns 
                st.session_state.original_columns = ( 
                    list( 
                        cleaned_df.columns 
                    ) 
                ) 
 
                # Start transformation stage 
                st.session_state.calculated_columns = [] 
 
                st.session_state.stage = ( 
                    "power_query_choice" 
                ) 
 
                st.rerun() 
 
 
# ============================================================ 
# STAGE 5 — TRANSFORMATION CHOICE 
# ============================================================ 
 
elif st.session_state.stage == "power_query_choice": 
 
    st.subheader( 
        "Transformation" 
    ) 
 
    st.write( 
        "Create or delete AI calculated columns, " 
        "or continue without transformation." 
    ) 
 
    if ( 
        st.session_state.transformed_df 
        is None 
    ): 
 
        st.session_state.transformed_df = ( 
            st.session_state.cleaned_df.copy() 
        ) 
 
    df = ( 
        st.session_state.transformed_df 
    ) 
 
    # -------------------------------------------------------- 
    # CURRENT CALCULATED COLUMNS 
    # -------------------------------------------------------- 
 
    calculated_columns = ( 
        get_calculated_columns( 
            df 
        ) 
    ) 
 
    if calculated_columns: 
 
        st.info( 
            "AI calculated columns currently available: " 
            + 
            ", ".join( 
                map( 
                    str, 
                    calculated_columns 
                ) 
            ) 
        ) 
 
    else: 
 
        st.caption( 
            "No AI calculated columns have " 
            "been created yet." 
        ) 
 
    # -------------------------------------------------------- 
    # OPEN CALCULATED COLUMN COMMAND SCREEN 
    # -------------------------------------------------------- 
 
    if st.button( 
        "Create / Delete Calculated Column", 
        type="primary" 
    ): 
 
        st.session_state.stage = ( 
            "power_query" 
        ) 
 
        st.rerun() 
 
    st.divider() 
 
    if st.button( 
        "Skip Transformation → Pivot" 
    ): 
 
        st.session_state.stage = ( 
            "pivot" 
        ) 
 
        st.rerun() 
 
 
# ============================================================ 
# STAGE 5B — CALCULATED COLUMN COMMAND 
# ============================================================ 
 
elif st.session_state.stage == "power_query": 
 
    st.subheader( 
        "Calculated Column Commands" 
    ) 
 
    df = ( 
        st.session_state.transformed_df 
    ) 
 
    st.write( 
        "Use one command at a time." 
    ) 
 
    st.write( 
        "Create:" 
    ) 
 
    st.code( 
        "Create a calculated column Revenue = Quantity * Price" 
    ) 
 
    st.write( 
        "Delete:" 
    ) 
 
    st.code( 
        "delete revenue" 
    ) 
 
    st.write( 
        "Other supported examples:" 
    ) 
 
    st.code( 
        "Create calculated column Profit = Revenue - Cost\n" 
        "Create calculated column Discount = Price * 0.90\n" 
        "delete profit\n" 
        "remove discount\n" 
        "delete calculated column Revenue" 
    ) 
 
    # -------------------------------------------------------- 
    # CURRENT CALCULATED COLUMNS 
    # -------------------------------------------------------- 
 
    calculated_columns = ( 
        get_calculated_columns( 
            df 
        ) 
    ) 
 
    if calculated_columns: 
 
        st.info( 
            "Current AI calculated columns: " 
            + 
            ", ".join( 
                map( 
                    str, 
                    calculated_columns 
                ) 
            ) 
        ) 
 
    # -------------------------------------------------------- 
    # COMMAND 
    # -------------------------------------------------------- 
 
    request = st.text_input( 
        "Enter your command" 
    ) 
 
    if st.button( 
        "Run Command", 
        type="primary" 
    ): 
 
        if not request.strip(): 
 
            st.warning( 
                "Please enter a command." 
            ) 
 
        # ---------------------------------------------------- 
        # DELETE 
        # ---------------------------------------------------- 
 
        elif is_delete_column_request( 
            request 
        ): 
 
            requested_column = ( 
                extract_delete_column_request( 
                    request 
                ) 
            ) 
 
            if requested_column is None: 
 
                st.error( 
                    "Please specify the calculated " 
                    "column name." 
                ) 
 
            else: 
 
                result, message = ( 
                    remove_calculated_column( 
                        df, 
                        requested_column 
                    ) 
                ) 
 
                if result is None: 
 
                    st.error( 
                        message 
                    ) 
 
                else: 
 
                    st.session_state.transformed_df = ( 
                        result 
                    ) 
 
                    st.session_state.transformation_log.append( 
                        message 
                    ) 
 
                    st.success( 
                        message 
                    ) 
 
                    st.dataframe( 
                        result.head(20), 
                        use_container_width=True, 
                        hide_index=True 
                    ) 
 
        # ---------------------------------------------------- 
        # CREATE 
        # ---------------------------------------------------- 
 
        else: 
 
            result, new_column, message = ( 
                create_calculated_column( 
                    df, 
                    request 
                ) 
            ) 
 
            if result is None: 
 
                st.error( 
                    message 
                ) 
 
            else: 
 
                st.session_state.transformed_df = ( 
                    result 
                ) 
 
                st.session_state.transformation_log.append( 
                    message 
                ) 
 
                st.success( 
                    message 
                ) 
 
                st.dataframe( 
                    result.head(20), 
                    use_container_width=True, 
                    hide_index=True 
                ) 
 
    st.divider() 
 
    if st.button( 
        "Transformation Complete → Pivot" 
    ): 
 
        st.session_state.stage = ( 
            "pivot" 
        ) 
 
        st.rerun() 
 
 
# ============================================================ 
# STAGE 6 — PIVOT 
# ============================================================ 
 
elif st.session_state.stage == "pivot": 
 
    st.subheader( 
        "Pivot Analysis" 
    ) 
 
    df = get_working_df() 
 
    if df is None: 
 
        st.error( 
            "No dataset available." 
        ) 
 
    else: 
 
        pivot = ( 
            create_pivot_table( 
                df 
            ) 
        ) 
 
        st.session_state.pivot_df = ( 
            pivot 
        ) 
 
        if pivot.empty: 
 
            st.info( 
                "A pivot table could not be automatically " 
                "created because the dataset does not contain " 
                "both categorical and numeric columns." 
            ) 
 
        else: 
 
            st.dataframe( 
                pivot, 
                use_container_width=True, 
                hide_index=True 
            ) 
 
        if st.button( 
            "Continue to Analysis", 
            type="primary" 
        ): 
 
            st.session_state.stage = ( 
                "analysis" 
            ) 
 
            st.rerun() 
 
 
# ============================================================ 
# STAGE 7 — ANALYSIS 
# ============================================================ 
 
elif st.session_state.stage == "analysis": 
 
    st.subheader( 
        "💬 Analyze Your Data" 
    ) 
 
    df = get_working_df() 
 
    if df is None: 
 
        st.error( 
            "No dataset available." 
        ) 
 
    else: 
 
        # ---------------------------------------------------- 
        # CHAT HISTORY 
        # ---------------------------------------------------- 
 
        for message in ( 
            st.session_state.chat_messages 
        ): 
 
            if message["role"] == "user": 
 
                st.chat_message( 
                    "user" 
                ).write( 
                    message["content"] 
                ) 
 
            else: 
 
                st.chat_message( 
                    "assistant" 
                ).write( 
                    message["content"] 
                ) 
 
        # ---------------------------------------------------- 
        # CHAT INPUT 
        # ---------------------------------------------------- 
 
        request = st.chat_input( 
            "Ask a question or give a command..." 
        ) 
 
        if request: 
 
            add_message( 
                "user", 
                request 
            ) 
 
            lower = ( 
                request.lower() 
                .strip() 
            ) 
 
            result = None 
 
            response = None 
 
            # =================================================
            # CREATE VISUALIZATION FROM AI CHAT
            # =================================================

            visualization_words = [
                "chart", "graph", "plot", "visualization", "visualisation",
                "bar chart", "line chart", "pie chart", "scatter", "histogram"
            ]

            if any(word in lower for word in visualization_words):
                fig, chart_data, chart_name, error = create_visualization_from_request(
                    df, request
                )

                if error:
                    response = error
                else:
                    chart_type = infer_chart_request(df, request)["type"]
                    add_visualization(
                        fig,
                        chart_data,
                        chart_name,
                        request,
                        chart_type
                    )
                    plt.close(fig)
                    response = (
                        f"Created '{chart_name}' and added it to the visualization area. "
                        "You can keep adding visualizations or delete this one independently."
                    )

            # =================================================
            # DELETE CALCULATED COLUMN
            # =================================================

            elif (business_answer := answer_business_question(df, request)) is not None:
                response = business_answer

            elif is_delete_column_request(request):
 
 
                requested_column = ( 
                    extract_delete_column_request( 
                        request 
                    ) 
                ) 
 
                if requested_column is None: 
 
                    response = ( 
                        "Please specify the calculated " 
                        "column you want to delete. " 
                        "Example: delete revenue." 
                    ) 
 
                else: 
 
                    updated_df, message = ( 
                        remove_calculated_column( 
                            df, 
                            requested_column 
                        ) 
                    ) 
 
                    if updated_df is None: 
 
                        response = message 
 
                    else: 
 
                        st.session_state.transformed_df = ( 
                            updated_df 
                        ) 
 
                        st.session_state.transformation_log.append( 
                            message 
                        ) 
 
                        df = updated_df 
 
                        response = message 
 
                        result = pd.DataFrame({ 
                            "Action": [ 
                                "Delete calculated column" 
                            ], 
                            "Column": [ 
                                requested_column 
                            ], 
                            "Status": [ 
                                "Deleted" 
                            ] 
                        }) 
 
            # ================================================= 
            # CREATE CALCULATED COLUMN 
            # ================================================= 
 
            elif ( 
                "=" in request 
            ): 
 
                result_df, new_column, message = ( 
                    create_calculated_column( 
                        df, 
                        request 
                    ) 
                ) 
 
                if result_df is None: 
 
                    response = message 
 
                else: 
 
                    st.session_state.transformed_df = ( 
                        result_df 
                    ) 
 
                    st.session_state.transformation_log.append( 
                        message 
                    ) 
 
                    df = result_df 
 
                    response = ( 
                        f"{message} " 
                        f"The new column is now available " 
                        f"in the transformed dataset." 
                    ) 
 
                    result = pd.DataFrame({ 
                        "Action": [ 
                            "Create calculated column" 
                        ], 
                        "Column": [ 
                            new_column 
                        ], 
                        "Status": [ 
                            "Created" 
                        ] 
                    }) 
 
            # ================================================= 
            # ROW COUNT 
            # ================================================= 
 
            elif ( 
                "how many rows" in lower 
                or 
                "row count" in lower 
            ): 
 
                result = pd.DataFrame({ 
                    "Metric": [ 
                        "Rows" 
                    ], 
                    "Value": [ 
                        len(df) 
                    ] 
                }) 
 
                response = ( 
                    f"The dataset contains " 
                    f"{len(df):,} rows." 
                ) 
 
            # ================================================= 
            # COLUMN COUNT 
            # ================================================= 
 
            elif ( 
                "how many columns" in lower 
                or 
                "column count" in lower 
            ): 
 
                result = pd.DataFrame({ 
                    "Metric": [ 
                        "Columns" 
                    ], 
                    "Value": [ 
                        len(df.columns) 
                    ] 
                }) 
 
                response = ( 
                    f"The dataset contains " 
                    f"{len(df.columns):,} columns." 
                ) 
 
            # ================================================= 
            # MISSING 
            # ================================================= 
 
            elif "missing" in lower: 
 
                missing = ( 
                    df.isna() 
                    .sum() 
                    .sort_values( 
                        ascending=False 
                    ) 
                ) 
 
                result = ( 
                    missing 
                    .reset_index() 
                ) 
 
                result.columns = [ 
                    "Column", 
                    "Missing Values" 
                ] 
 
                response = ( 
                    "Here is the missing-value summary." 
                ) 
 
            # ================================================= 
            # DUPLICATES 
            # ================================================= 
 
            elif "duplicate" in lower: 
 
                duplicates = int( 
                    df.duplicated().sum() 
                ) 
 
                result = pd.DataFrame({ 
                    "Metric": [ 
                        "Duplicate Rows" 
                    ], 
                    "Value": [ 
                        duplicates 
                    ] 
                }) 
 
                response = ( 
                    f"The dataset contains " 
                    f"{duplicates:,} duplicate row(s)." 
                ) 
 
            # ================================================= 
            # CALCULATED COLUMNS 
            # ================================================= 
 
            elif ( 
                "calculated column" in lower 
                or 
                "calculated columns" in lower 
            ): 
 
                calculated_columns = ( 
                    get_calculated_columns( 
                        df 
                    ) 
                ) 
 
                if calculated_columns: 
 
                    result = pd.DataFrame({ 
                        "Calculated Columns": 
                            calculated_columns 
                    }) 
 
                    response = ( 
                        "These are the calculated " 
                        "columns created by the AI." 
                    ) 
 
                else: 
 
                    response = ( 
                        "No AI calculated columns " 
                        "currently exist." 
                    ) 
 
            # ================================================= 
            # NUMERIC SUMMARY 
            # ================================================= 
 
            elif ( 
                "numeric summary" in lower 
                or 
                "statistics" in lower 
                or 
                "describe" in lower 
            ): 
 
                result = ( 
                    numeric_summary( 
                        df 
                    ) 
                ) 
 
                response = ( 
                    "Here is the numeric summary." 
                ) 
 
            # ================================================= 
            # TOP 
            # ================================================= 
 
            elif "top" in lower: 
 
                categorical = [ 
                    col 
                    for col in df.columns 
                    if ( 
                        df[col].dtype == "object" 
                        or 
                        pd.api.types.is_string_dtype( 
                            df[col] 
                        ) 
                    ) 
                ] 
 
                if categorical: 
 
                    column = ( 
                        categorical[0] 
                    ) 
 
                    result = ( 
                        top_categories( 
                            df, 
                            column, 
                            5 
                        ) 
                    ) 
 
                    response = ( 
                        f"Top values for " 
                        f"'{column}'." 
                    ) 
 
                else: 
 
                    response = ( 
                        "No categorical column was " 
                        "found for a top-value analysis." 
                    ) 
 
            # ================================================= 
            # REVENUE / NUMERIC 
            # ================================================= 
 
            elif any( 
                word in lower 
                for word in [ 
                    "revenue", 
                    "sales", 
                    "amount", 
                    "total" 
                ] 
            ): 
 
                analysis = ( 
                    revenue_analysis( 
                        df 
                    ) 
                ) 
 
                if analysis: 
 
                    result = pd.DataFrame({ 
                        "Metric": [ 
                            "Column", 
                            "Total", 
                            "Average", 
                            "Minimum", 
                            "Maximum" 
                        ], 
 
                        "Value": [ 
                            analysis["column"], 
                            analysis["sum"], 
                            analysis["average"], 
                            analysis["minimum"], 
                            analysis["maximum"] 
                        ] 
                    }) 
 
                    response = ( 
                        f"Using " 
                        f"'{analysis['column']}' " 
                        f"as the relevant numeric " 
                        f"measure, the total is " 
                        f"{analysis['sum']:,.2f}." 
                    ) 
 
                else: 
 
                    response = ( 
                        "No numeric columns were found." 
                    ) 
 
            # ================================================= 
            # DEFAULT 
            # ================================================= 
 
            else: 
 
                response = ( 
                    "I can analyze row counts, columns, " 
                    "missing values, duplicates, numeric " 
                    "statistics, top categories, totals, " 
                    "charts, and calculated columns. " 
                    "For example: " 
                    "'Create a calculated column " 
                    "Revenue = Quantity * Price' " 
                    "or 'Show a horizontal bar chart of revenue by hotel'." 
                ) 
 
            # ------------------------------------------------ 
            # ASSISTANT RESPONSE 
            # ------------------------------------------------ 
 
            add_message( 
                "assistant", 
                response 
            ) 
 
            # ------------------------------------------------ 
            # STORE RESULT 
            # ------------------------------------------------ 
 
            if ( 
                isinstance( 
                    result, 
                    pd.DataFrame 
                ) 
                and 
                not result.empty 
            ): 
 
                st.session_state.analysis_results.append( 
                    result 
                ) 
 
                st.session_state.analysis_log.append( 
                    request 
                ) 
 
                st.session_state.current_result = ( 
                    result 
                ) 
 
            st.rerun() 
 
        # ---------------------------------------------------- 
        # CURRENT CALCULATED COLUMNS 
        # ---------------------------------------------------- 
 
        calculated_columns = ( 
            get_calculated_columns( 
                df 
            ) 
        ) 
 
        if calculated_columns: 
 
            with st.expander( 
                "AI Calculated Columns" 
            ): 
 
                st.dataframe( 
                    pd.DataFrame({ 
                        "Calculated Column": 
                            calculated_columns 
                    }), 
                    use_container_width=True, 
                    hide_index=True 
                ) 
 
        # ====================================================
        # VISUALIZATION COMMANDS + MANAGER
        # ====================================================

        # Visualizations can be created directly from the AI chat.
        # Example: "Show a horizontal bar chart of revenue by hotel"
        # The chart intent is handled before generic revenue analysis,
        # so "revenue" does not get misclassified as a KPI request.

        render_visualization_manager()

        st.divider()
        st.subheader("➕ Add Visualization")

        chart_request = st.text_input(
            "Describe the visualization you want",
            placeholder="Show a horizontal bar chart of revenue by hotel",
            key="visualization_request"
        )

        if st.button("Add Visualization", type="secondary"):
            if not chart_request.strip():
                st.warning("Please describe the visualization.")
            else:
                fig, chart_data, chart_name, error = create_visualization_from_request(
                    df, chart_request
                )

                if error:
                    st.error(error)
                else:
                    add_visualization(
                        fig,
                        chart_data,
                        chart_name,
                        chart_request,
                        infer_chart_request(df, chart_request)["type"]
                    )
                    plt.close(fig)
                    st.success(f"Added visualization: {chart_name}")
                    st.rerun()

        st.divider() 
 
        if st.button( 
            "Continue to Report", 
            type="primary" 
        ): 
 
            st.session_state.stage = ( 
                "report" 
            ) 
 
            st.rerun() 
 
 
# ============================================================ 
# STAGE 8 — REPORT 
# ============================================================ 
 
elif st.session_state.stage == "report": 
 
    st.subheader( 
        "📄 Excel Report" 
    ) 
 
    st.write( 
        "Your Excel report will contain the " 
        "raw data, cleaned data, transformed data, " 
        "validation, analysis and visualizations." 
    ) 
 
    # -------------------------------------------------------- 
    # SUMMARY 
    # -------------------------------------------------------- 
 
    df = get_working_df() 
 
    if df is not None: 
 
        col1, col2, col3 = ( 
            st.columns(3) 
        ) 
 
        with col1: 
 
            st.metric( 
                "Final Rows", 
                len(df) 
            ) 
 
        with col2: 
 
            st.metric( 
                "Final Columns", 
                len(df.columns) 
            ) 
 
        with col3: 
 
            st.metric( 
                "Missing Cells", 
                int( 
                    df.isna() 
                    .sum() 
                    .sum() 
                ) 
            ) 
 
    # -------------------------------------------------------- 
    # REPORT CONTENT 
    # -------------------------------------------------------- 
 
    st.markdown( 
        """ 
        ### Report Contents 
 
        - Raw Data 
        - Cleaned Data 
        - Transformed Data 
        - Pivot Table 
        - Cleaning Validation 
        - Cleaning Log 
        - Transformation Log 
        - Calculated Columns 
        - Analysis Log 
        - Analysis Results 
        - Business Insights 
        - Visualizations 
        """ 
    ) 
 
    if st.button( 
        "Generate Excel Report", 
        type="primary" 
    ): 
 
        try: 
 
            report_bytes = ( 
                generate_excel_report() 
            ) 
 
            st.session_state.report_generated = ( 
                True 
            ) 
 
            st.download_button( 
                label="⬇️ Download Excel Report", 
                data=report_bytes, 
                file_name=( 
                    "AI_Excel_Analyst_Report.xlsx" 
                ), 
                mime=( 
                    "application/vnd.openxmlformats-officedocument." 
                    "spreadsheetml.sheet" 
                ) 
            ) 
 
            st.success( 
                "Excel report generated successfully." 
            ) 
 
        except Exception as e: 
 
            st.error( 
                f"Unable to generate report: {e}" 
            ) 
 
    st.divider() 
 
    if st.button( 
        "Start New Analysis" 
    ): 
 
        for key, value in ( 
            DEFAULT_STATE.items() 
        ): 
 
            st.session_state[key] = value 
 
        st.rerun() 