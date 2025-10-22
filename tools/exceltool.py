import json
import textwrap
import traceback
from typing import Any, Dict
import pandas as pd
from agno.tools import tool
import google.generativeai as genai

# Configure Gemini API

GEMINI_API_KEY=""
genai.configure(api_key=GEMINI_API_KEY)

# Define the core function without the decorator
def load_excel(path: str) -> Dict[str, Any]:
    """
    Loads an Excel file and returns a dictionary containing:
      - 'df': DataFrame
      - 'columns': list of column names
      - 'preview': preview of the first 5 rows
      - 'shape': tuple representing the shape of the DataFrame
      - 'path': path to the Excel file
    """
    if not path:
        raise ValueError("A path to the Excel file is required.")
    df = pd.read_excel(path)
    return {
        "df": df,
        "columns": list(df.columns),
        "preview": df.head(5).to_dict(orient="records"),
        "shape": df.shape,
        "path": path,
    }

# Define the exceltool with the @tool decorator wrapping the core function
@tool(
    name="exceltool",
    description="Loads an Excel file and returns its metadata and preview.",
    show_result=True,
    stop_after_tool_call=True
)
def exceltool(path: str) -> Dict[str, Any]:
    """Tool wrapper for loading Excel files."""
    return load_excel(path)

# Helper function to find column name case-insensitively with fuzzy matching
def find_column(df: pd.DataFrame, search_term: str) -> str:
    """
    Finds a column in the DataFrame that matches the search term (case-insensitive).
    Also supports partial matching (e.g., 'price' matches 'Unit_Price').
    Returns the actual column name or None if not found.
    """
    search_lower = search_term.lower().replace('_', '').replace(' ', '')
    
    # First try exact match (ignoring case)
    for col in df.columns:
        col_normalized = str(col).lower().replace('_', '').replace(' ', '')
        if col_normalized == search_lower:
            return col
    
    # Then try partial match (search term is contained in column name)
    for col in df.columns:
        col_normalized = str(col).lower().replace('_', '').replace(' ', '')
        if search_lower in col_normalized:
            return col
    
    return None

# Function to execute the code snippet
def execute_snippet(df: pd.DataFrame, code: str) -> Dict[str, Any]:
    """
    Executes the provided code snippet in a restricted environment.
    Returns a dictionary with the execution result.
    """
    # Prepare the local environment with the DataFrame
    local_vars = {"df": df.copy(), "pd": pd, "find_column": find_column}
    try:
        # Execute the code snippet
        exec(textwrap.dedent(code), {}, local_vars)
        result = local_vars.get("result", None)
        if isinstance(result, pd.DataFrame):
            result = result.head(50).to_dict(orient="records")
        return {"ok": True, "result": result}
    except Exception as e:
        tb = traceback.format_exc()
        return {"ok": False, "error": str(e), "traceback": tb}

# Function to handle the user's query and process the Excel file
def answer_excel_question(path: str, user_query: str) -> Dict[str, Any]:
    """
    Processes the user's query by loading the Excel file, generating code to
    apply the requested operation, executing the code, and returning the results.
    """
    # Load the Excel file using the core load_excel function
    tool_out = load_excel(path)
    df = tool_out["df"]
    columns = tool_out["columns"]
    preview = tool_out["preview"]

    # Generate code based on the user's query
    code = generate_code(columns, preview, user_query)

    # Execute the generated code
    exec_out = execute_snippet(df, code)

    return {
        "path": path,
        "columns": columns,
        "preview": preview,
        "query": user_query,
        "code": code,
        "execution": exec_out,
    }

# Function to generate code based on the user's query using Gemini
def generate_code(columns: list, preview: list, user_query: str) -> str:
    """
    Generates a Python code snippet based on the user's query, the columns of the
    DataFrame, and a preview of the data using Gemini AI.
    """
    
    prompt = f"""You are a Python pandas expert. Generate Python code to answer the user's query about an Excel DataFrame.

**Available DataFrame columns:**
{json.dumps(columns, indent=2)}

**Preview of data (first 5 rows):**
{json.dumps(preview, indent=2)}

**User Query:**
{user_query}

**Critical Instructions:**
1. The DataFrame is available as variable 'df'
2. MANDATORY: Use 'find_column(df, "column_name")' to find ANY column before using it
   - Example: item_col = find_column(df, "item") 
   - Example: qty_col = find_column(df, "quantity")
   - This handles case-insensitive matching and variations
3. ALWAYS check if column is None before using it
4. ALWAYS store the final result in a variable called 'result'
5. Analyze the query carefully:
   - "total available items" = count unique items OR sum quantities
   - "no of glue available" = filter where Item='Glue', then sum Quantity or count rows
   - "items with no discount" = filter where Discount(%)=0
   - "apply discount" = modify price values
6. For numeric operations: pd.to_numeric(df[col], errors='coerce')
7. Work on copy: df = df.copy()
8. For simple counts/sums, result can be a number or dict
9. For filtered data, convert DataFrame to dict: result = filtered_df.to_dict(orient='records')

**Example Code Patterns:**

# For "no of glue available" or "count glue items":
item_col = find_column(df, "item")
qty_col = find_column(df, "quantity")
if item_col and qty_col:
    df[qty_col] = pd.to_numeric(df[qty_col], errors='coerce')
    glue_rows = df[df[item_col].str.lower().str.contains('glue', na=False)]
    result = {{"total_quantity": int(glue_rows[qty_col].sum()), "item_count": len(glue_rows)}}

# For "total available items":
item_col = find_column(df, "item")
if item_col:
    result = {{"unique_items": df[item_col].nunique(), "total_rows": len(df)}}

**Generate ONLY the Python code, no explanations or markdown:**"""

    try:
        model = genai.GenerativeModel('gemini-2.5-flash')
        response = model.generate_content(prompt)
        code = response.text.strip()
        
        # Clean up the code (remove markdown code blocks if present)
        if code.startswith('```python'):
            code = code[len('```python'):].strip()
        if code.startswith('```'):
            code = code[3:].strip()
        if code.endswith('```'):
            code = code[:-3].strip()
            
        return code
    except Exception as e:
        # Fallback to a simple code template if Gemini fails
        print(f"Warning: Gemini API failed ({e}), using fallback code generation")
        return """
# Fallback: Display all data
result = df.head(50).to_dict(orient='records')
"""