import os
import json
import textwrap
import traceback
from datetime import datetime
from typing import Any, Dict, Union
import pandas as pd
from agno.tools import tool
import google.generativeai as genai

# ===============================
# Gemini API Configuration
# ===============================
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
genai.configure(api_key=GEMINI_API_KEY)

# ===============================
# Load Excel or CSV Data
# ===============================
def load_excel_or_csv(path: str, sheet_name: Union[str, None] = None) -> Dict[str, Any]:
    if not path:
        raise ValueError("A valid path to the Excel or CSV file is required.")
    if not os.path.exists(path):
        raise FileNotFoundError(f"File not found: {path}")
    ext = os.path.splitext(path)[1].lower()
    df = None
    active_sheet = None
    if ext in [".xls", ".xlsx"]:
        excel_file = pd.ExcelFile(path)
        sheets = excel_file.sheet_names
        if len(sheets) > 1:
            print(f"\n📘 This Excel file contains multiple sheets: {sheets}")
            if sheet_name is None:
                sheet_name = input("Enter sheet name to use (default first): ").strip() or sheets[0]
        else:
            sheet_name = sheets[0]
        df = pd.read_excel(path, sheet_name=sheet_name)
        active_sheet = sheet_name
    elif ext == ".csv":
        df = pd.read_csv(path)
        active_sheet = "CSV"
    else:
        raise ValueError("Unsupported file type. Only .xlsx, .xls, and .csv are supported.")
    return {
        "df": df,
        "columns": list(df.columns),
        "preview": df.head(5).to_dict(orient="records"),
        "shape": df.shape,
        "path": path,
        "sheet_name": active_sheet,
    }

# ===============================
# Agno Tool Wrapper
# ===============================
@tool(
    name="exceltool",
    description="Loads an Excel (multi-sheet) or CSV file and returns metadata and preview.",
    show_result=True,
    stop_after_tool_call=True
)
def exceltool(path: str, sheet_name: Union[str, None] = None) -> Dict[str, Any]:
    return load_excel_or_csv(path, sheet_name)

# ===============================
# Helper Function: Find Column
# ===============================
def find_column(df: pd.DataFrame, search_term: str) -> Union[str, None]:
    search_lower = search_term.lower().replace("_", "").replace(" ", "")
    for col in df.columns:
        normalized = str(col).lower().replace("_", "").replace(" ", "")
        if normalized == search_lower:
            return col
    for col in df.columns:
        normalized = str(col).lower().replace("_", "").replace(" ", "")
        if search_lower in normalized:
            return col
    return None

# ===============================
# Execute Generated Code Safely
# ===============================
def execute_snippet(df: pd.DataFrame, code: str) -> Dict[str, Any]:
    local_vars = {"df": df.copy(), "pd": pd, "find_column": find_column}
    try:
        exec(textwrap.dedent(code), {}, local_vars)
        result = local_vars.get("result", None)
        if isinstance(result, pd.DataFrame):
            # convert for result output
            result = result.head(50).to_dict(orient="records")
        return {"ok": True, "result": result}
    except Exception as e:
        return {"ok": False, "error": str(e), "traceback": traceback.format_exc()}

# ===============================
# Generate Pandas Code via Gemini
# ===============================
def generate_code(columns: list, preview: list, user_query: str, path: str, sheet_name: str) -> str:
    prompt = f"""
You are a Python pandas expert. Generate Python code to answer the user's query about an Excel or CSV DataFrame.

**File Information:**
- Path: {path}
- Sheet Name: {sheet_name}

**Available DataFrame columns:**
{json.dumps(columns, indent=2)}

**Preview of data (first 5 rows):**
{json.dumps(preview, indent=2)}

**User Query:**
{user_query}

**Critical Instructions:**
1. The DataFrame is available as variable 'df'.
2. ALWAYS use find_column(df, "column_name") to reference columns (case-insensitive).
3. ALWAYS check if column is None before using.
4. Store final output in variable 'result'.
5. Use pd.to_numeric(df[col], errors='coerce') for numeric ops.
6. Return result as:
   - Dict for summary numbers.
   - List of dicts for filtered rows.
7 if remaining fields are not given by the user just simply return  0 for number fields ( if no inter dependency  , if there is inter dependency then  calculate it and make value  for example user given cost of item , and quantity of item and total is not mention so you are responsible to calculate , similarly for other fields also  ) and "" for string fields for there remaining fields 
8 Model is not hallucinating but get real data from the excel file or user query.

**Generate ONLY executable Python code. No markdown, no explanations.**
"""
    model = genai.GenerativeModel("gemini-2.5-flash")
    response = model.generate_content(prompt)
    code = response.text.strip()
    for marker in ("```python", "```"):
        code = code.replace(marker, "")
    return code.strip()

# ===============================
# Save Query Results to JSON
# ===============================
def save_query_output(query: str, code: str, result: Any) -> None:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    results_dir = "results"
    os.makedirs(results_dir, exist_ok=True)
    filename = os.path.join(results_dir, f"query_result_{timestamp}.json")
    payload = {
        "timestamp": timestamp,
        "query": query,
        "generated_code": code,
        "result": result,
    }
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, default=str)
    print(f"\n💾 Query and results saved to: {filename}")

# ===============================
# Save Query Results to Excel (new sheet/file)
# ===============================
def save_result_to_excel(result: Any, base_path: str, query: str) -> None:
    """
    Saves the result (dict/list or single value) into an Excel file.
    If result is list of dicts, convert to DataFrame.
    If result is a simple number or dict, wrap into one-row DataFrame.
    The new file is named based on base_path and query timestamp.
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    # derive filename
    dirname, fname = os.path.split(base_path)
    name, ext = os.path.splitext(fname)
    new_fname = f"{name}_result_{timestamp}.xlsx"
    new_path =  os.path.join(dirname, new_fname)
    # convert result to DataFrame
    if isinstance(result, list):
        df_result = pd.DataFrame(result)
    elif isinstance(result, dict):
        df_result = pd.DataFrame([result])
    else:
        df_result = pd.DataFrame([{"result": result}])
    # write to excel
    with pd.ExcelWriter(new_path) as writer:
        df_result.to_excel(writer, sheet_name="QueryResult", index=False)
    print(f"📄 Result saved to Excel file: {new_path}")

# ===============================
# Main Logic: Answer Excel Question
# ===============================
def answer_excel_question(path: str, user_query: str, sheet_name: Union[str, None] = None) -> Dict[str, Any]:
    tool_out = load_excel_or_csv(path, sheet_name)
    df = tool_out["df"]
    columns = tool_out["columns"]
    preview = tool_out["preview"]
    sheet = tool_out["sheet_name"]
    # generate code
    code = generate_code(columns, preview, user_query, path, sheet)
    # execute snippet
    exec_out = execute_snippet(df, code)
    return {
        "path": path,
        "sheet": sheet,
        "columns": columns,
        "query": user_query,
        "code": code,
        "execution": exec_out,
    }

# ===============================
# Interactive Loop Tool
# ===============================
@tool(
    name="interactive_loop",
    description="Interactive loop to handle user input.",
    show_result=True
)
def interactive_loop(path: str, user_query: str) -> None:
    print("=== Agno + Gemini Excel/CSV Agent ===")
    path = path.strip()
    if not path:
        print("❌ No path provided; exiting.")
        return
    print(f"✅ Loaded file: {path}")
    user_query = user_query.strip()
    if user_query.lower() in {"exit", "quit"}:
        print("👋 Goodbye.")
        return
    print("\n🤖 Generating code with Gemini...")
    out = answer_excel_question(path, user_query)
    print("\n--- Generated Code ---")
    print(out["code"])
    print("\n--- Execution Result ---")
    exec_out = out["execution"]
    if exec_out.get("ok"):
        result = exec_out["result"]
        print(json.dumps(result, indent=2, default=str))
        # Save JSON?

        save_query_output(user_query, out["code"], result)
        save_result_to_excel(result, path, user_query)
    else:
        print(f"❌ Error: {exec_out['error']}")
        print(exec_out.get("traceback", ""))
 
 
if __name__ == "__main__":
    interactive_loop(
        path="/Users/jayanth/Desktop/SheetShift/Invoice_20rows.xlsx",
        user_query="""   can you create a new sheet i want Eraser , ruler , of quantity 69 each and each price is 5 and discount is 10%
""",
        save_choice="y"
    )
