import os
import json
import textwrap
import traceback
from datetime import datetime
from typing import Any, Dict, Union
import pandas as pd
from agno.tools import tool
from agno.agent import Agent
from agno.models.google import Gemini
import google.generativeai as genai

# ===============================
# Gemini API Configuration
# ===============================
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
genai.configure(api_key=GEMINI_API_KEY)

# ===============================
# Load All Sheets from Excel
# ===============================
def load_all_sheets(path: str) -> Dict[str, Any]:
    """Load all sheets from an Excel file and return metadata for each."""
    if not path:
        raise ValueError("A valid path to the Excel file is required.")
    if not os.path.exists(path):
        raise FileNotFoundError(f"File not found: {path}")
    
    ext = os.path.splitext(path)[1].lower()
    
    if ext == ".csv":
        df = pd.read_csv(path)
        return {
            "sheets": {
                "CSV": {
                    "df": df,
                    "columns": list(df.columns),
                    "preview": df.head(5).to_dict(orient="records"),
                    "shape": df.shape
                }
            },
            "path": path,
            "sheet_names": ["CSV"]
        }
    
    elif ext in [".xls", ".xlsx"]:
        excel_file = pd.ExcelFile(path)
        sheets_data = {}
        
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(path, sheet_name=sheet_name)
            sheets_data[sheet_name] = {
                "df": df,
                "columns": list(df.columns),
                "preview": df.head(5).to_dict(orient="records"),
                "shape": df.shape
            }
        
        print(f"\n📚 Loaded {len(excel_file.sheet_names)} sheets: {excel_file.sheet_names}")
        
        return {
            "sheets": sheets_data,
            "path": path,
            "sheet_names": excel_file.sheet_names
        }
    else:
        raise ValueError("Unsupported file type. Only .xlsx, .xls, and .csv are supported.")

# ===============================
# Helper Function: Find Column
# ===============================
def find_column(df: pd.DataFrame, search_term: str) -> Union[str, None]:
    """Find column in DataFrame with fuzzy matching."""
    search_lower = search_term.lower().replace("_", "").replace(" ", "")
    
    # Exact match
    for col in df.columns:
        normalized = str(col).lower().replace("_", "").replace(" ", "")
        if normalized == search_lower:
            return col
    
    # Partial match
    for col in df.columns:
        normalized = str(col).lower().replace("_", "").replace(" ", "")
        if search_lower in normalized or normalized in search_lower:
            return col
    
    return None

# ===============================
# Detect Query Intent
# ===============================
def detect_query_intent(user_query: str) -> Dict[str, Any]:
    """Use Gemini to understand if user wants to create new sheet or update existing ones."""
    prompt = f"""
Analyze this user query and determine the intent:

User Query: {user_query}

Respond ONLY with a JSON object (no markdown, no extra text):
{{
    "intent": "create_new_sheet" or "update_existing" or "query_data",
    "target_sheets": ["all"] or ["specific_sheet_name"] or null,
    "new_sheet_name": "name" or null,
    "fields_to_update": ["field1", "field2"] or null,
    "description": "brief description of what user wants"
}}

Rules:
- If query mentions "create new sheet" or "add new sheet", intent is "create_new_sheet"
- If query mentions updating/changing values or incrementing, intent is "update_existing"
- If query is asking questions about data, intent is "query_data"
- For updates, if user mentions "all sheets" or doesn't specify, set target_sheets to ["all"]
"""
    
    model = genai.GenerativeModel("gemini-2.5-flash")
    response = model.generate_content(prompt)
    
    try:
        result_text = response.text.strip()
        for marker in ("```json", "```"):
            result_text = result_text.replace(marker, "")
        intent_data = json.loads(result_text.strip())
        return intent_data
    except Exception as e:
        print(f"⚠️  Could not parse intent: {e}")
        return {
            "intent": "query_data",
            "target_sheets": None,
            "new_sheet_name": None,
            "fields_to_update": None,
            "description": user_query
        }

# ===============================
# Generate Code for New Sheet Creation
# ===============================
def generate_new_sheet_code(user_query: str, existing_sheets_info: Dict[str, Any], file_path: str) -> str:
    """Generate code to create a new sheet with user-specified data."""
    
    # Get column structure from existing sheets for reference
    sample_columns = {}
    for sheet_name, info in existing_sheets_info.items():
        sample_columns[sheet_name] = info["columns"]
        break  # Just get first sheet as reference
    
    prompt = f"""
You are a Python pandas expert. Generate code to create a NEW DataFrame for a new sheet based on user requirements.

**Existing Sheets in File:**
{json.dumps({name: info["columns"] for name, info in existing_sheets_info.items()}, indent=2)}

**User Query:**
{user_query}

**CRITICAL Instructions:**
1. Create a new DataFrame with columns based on user requirements
2. Parse the user query to extract: items, quantities, prices, discounts, etc.
3. Calculate derived fields (like Total = Price * Quantity, Final_Amount = Total * (1 - Discount))
4. Store the new DataFrame in variable 'new_df' (REQUIRED)
5. Store the new sheet name in variable 'new_sheet_name' (REQUIRED)
6. If user doesn't specify a sheet name, use "NewItems" or similar descriptive name
7. DO NOT try to read or write to files - ONLY create the DataFrame
8. DO NOT use pd.read_excel or pd.ExcelWriter in your code
9. For discount calculations, convert percentages to decimals (50% = 0.50)
10. Include columns like: Item, Quantity, Unit_Price, Discount(%), Total, Final_Amount

**Example Structure (DO NOT include file operations):**
```python
import pandas as pd

# Parse user data
items = ["Highlighter", "Stapler", "Ruler"]
quantities = [69, 69, 69]
unit_prices = [5, 5, 5]
discount_percentages = [50, 50, 50]

# Calculate derived fields
totals = [q * p for q, p in zip(quantities, unit_prices)]
final_amounts = [t * (1 - d/100) for t, d in zip(totals, discount_percentages)]

# Create DataFrame
new_df = pd.DataFrame({{
    "Item": items,
    "Quantity": quantities,
    "Unit_Price": unit_prices,
    "Discount(%)": discount_percentages,
    "Total": totals,
    "Final_Amount": final_amounts
}})

new_sheet_name = "StationeryItems"
```

Generate ONLY the DataFrame creation code. No file operations. No markdown, no explanations.
"""
    
    model = genai.GenerativeModel("gemini-2.5-flash")
    response = model.generate_content(prompt)
    code = response.text.strip()
    
    for marker in ("```python", "```"):
        code = code.replace(marker, "")
    
    return code.strip()

# ===============================
# Generate Code for Multi-Sheet Updates
# ===============================
def generate_update_code(columns: list, preview: list, user_query: str, sheet_name: str, all_sheets_info: Dict[str, Any]) -> str:
    """Generate code to update data across sheets."""
    prompt = f"""
You are a Python pandas expert. Generate code to UPDATE data in a DataFrame based on user query.

**Current Sheet:** {sheet_name}
**Current Columns:** {json.dumps(columns, indent=2)}
**Preview:** {json.dumps(preview, indent=2)}

**All Sheets Info:**
{json.dumps({name: info["columns"] for name, info in all_sheets_info.items()}, indent=2)}

**User Query:** {user_query}

**Instructions:**
1. DataFrame is available as 'df'
2. Use find_column(df, "column_name") for column access
3. Always check if column exists before using
4. For updates, modify df in place or create new result df
5. Store result in 'result' variable (the modified DataFrame)
6. Handle dependencies (e.g., if incrementing employee count, update related fields)
7. Use pd.to_numeric() for numeric operations
8. DO NOT use pd.read_excel or any file operations

**Generate ONLY executable Python code. No markdown, no explanations.**
"""
    
    model = genai.GenerativeModel("gemini-2.5-flash")
    response = model.generate_content(prompt)
    code = response.text.strip()
    
    for marker in ("```python", "```"):
        code = code.replace(marker, "")
    
    return code.strip()

# ===============================
# Generate Code for Query
# ===============================
def generate_query_code(columns: list, preview: list, user_query: str, sheet_name: str) -> str:
    """Generate code to query/analyze data."""
    prompt = f"""
You are a Python pandas expert. Generate code to answer user's query about data.

**Sheet:** {sheet_name}
**Columns:** {json.dumps(columns, indent=2)}
**Preview:** {json.dumps(preview, indent=2)}

**User Query:** {user_query}

**Instructions:**
1. DataFrame is 'df'
2. Use find_column(df, "column_name")
3. Check if column is None before using
4. Store result in 'result' variable
5. Result can be: dict, list of dicts, or DataFrame
6. Use pd.to_numeric() for numeric ops
7. DO NOT use file operations

Generate ONLY executable Python code. No markdown, no explanations.
"""
    
    model = genai.GenerativeModel("gemini-2.5-flash")
    response = model.generate_content(prompt)
    code = response.text.strip()
    
    for marker in ("```python", "```"):
        code = code.replace(marker, "")
    
    return code.strip()

# ===============================
# Execute Code Safely
# ===============================
def execute_snippet(df: pd.DataFrame, code: str, extra_vars: Dict[str, Any] = None) -> Dict[str, Any]:
    """Execute generated code with DataFrame."""
    local_vars = {"df": df.copy(), "pd": pd, "find_column": find_column, "datetime": datetime}
    
    if extra_vars:
        local_vars.update(extra_vars)
    
    try:
        exec(textwrap.dedent(code), {}, local_vars)
        result = local_vars.get("result", None)
        new_sheet_name = local_vars.get("new_sheet_name", None)
        new_df = local_vars.get("new_df", None)
        
        return {
            "ok": True,
            "result": result,
            "new_sheet_name": new_sheet_name,
            "new_df": new_df
        }
    except Exception as e:
        return {
            "ok": False,
            "error": str(e),
            "traceback": traceback.format_exc()
        }

# ===============================
# Save Updated Excel with All Sheets
# ===============================
def save_updated_excel(path: str, sheets_data: Dict[str, pd.DataFrame], new_sheets: Dict[str, pd.DataFrame] = None) -> str:
    """Save all sheets (original + updated + new) back to Excel."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dirname, fname = os.path.split(path)
    name, ext = os.path.splitext(fname)
    new_path = os.path.join(dirname, f"{name}_updated_{timestamp}.xlsx")
    
    try:
        with pd.ExcelWriter(new_path, engine='openpyxl') as writer:
            # Write all updated existing sheets
            for sheet_name, df in sheets_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Write new sheets if any
            if new_sheets:
                for sheet_name, df in new_sheets.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"\n💾 Updated Excel saved to: {new_path}")
        return new_path
    except Exception as e:
        print(f"\n❌ Error saving file: {e}")
        raise

# ===============================
# Process User Query Across All Sheets
# ===============================
def process_query_all_sheets(path: str, user_query: str) -> Dict[str, Any]:
    """Main function to process user query across all sheets."""
    
    # Load all sheets
    all_data = load_all_sheets(path)
    sheets = all_data["sheets"]
    sheet_names = all_data["sheet_names"]
    
    # Detect intent
    intent_info = detect_query_intent(user_query)
    print(f"\n🎯 Detected Intent: {intent_info['intent']}")
    print(f"📋 Description: {intent_info['description']}")
    
    results = {
        "intent": intent_info["intent"],
        "sheets_processed": [],
        "new_sheets": {},
        "errors": []
    }
    
    # Handle based on intent
    if intent_info["intent"] == "create_new_sheet":
        print("\n🆕 Creating new sheet...")
        code = generate_new_sheet_code(user_query, sheets, path)
        print("\n--- Generated Code ---")
        print(code)
        
        exec_result = execute_snippet(pd.DataFrame(), code)
        
        if exec_result["ok"] and exec_result["new_df"] is not None:
            new_sheet_name = exec_result["new_sheet_name"] or f"NewSheet_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            results["new_sheets"][new_sheet_name] = exec_result["new_df"]
            print(f"\n✅ New sheet '{new_sheet_name}' created with {len(exec_result['new_df'])} rows")
            print("\n📊 Preview of new sheet:")
            print(exec_result["new_df"].to_string(index=False))
        else:
            error_msg = f"Failed to create new sheet: {exec_result.get('error', 'Unknown error')}"
            results["errors"].append(error_msg)
            print(f"\n❌ Error: {exec_result.get('error')}")
            if exec_result.get("traceback"):
                print(exec_result["traceback"])
    
    elif intent_info["intent"] == "update_existing":
        print("\n🔄 Updating existing sheets...")
        target_sheets = intent_info.get("target_sheets", ["all"])
        
        sheets_to_update = sheet_names if "all" in target_sheets else target_sheets
        
        for sheet_name in sheets_to_update:
            if sheet_name not in sheets:
                continue
            
            sheet_info = sheets[sheet_name]
            df = sheet_info["df"]
            columns = sheet_info["columns"]
            preview = sheet_info["preview"]
            
            print(f"\n📝 Processing sheet: {sheet_name}")
            code = generate_update_code(columns, preview, user_query, sheet_name, sheets)
            print(f"\n--- Generated Code for {sheet_name} ---")
            print(code)
            
            exec_result = execute_snippet(df, code)
            
            if exec_result["ok"]:
                updated_df = exec_result["result"]
                if isinstance(updated_df, pd.DataFrame):
                    sheets[sheet_name]["df"] = updated_df
                    results["sheets_processed"].append(sheet_name)
                    print(f"✅ Updated sheet '{sheet_name}'")
                    print(f"📊 Preview of updated data:")
                    print(updated_df.head().to_string(index=False))
                else:
                    print(f"⚠️  No DataFrame returned for {sheet_name}")
            else:
                error_msg = f"Error in {sheet_name}: {exec_result.get('error')}"
                results["errors"].append(error_msg)
                print(f"❌ {error_msg}")
    
    else:  # query_data
        print("\n🔍 Querying data...")
        # Use first sheet or specified sheet
        target_sheet = sheet_names[0] if sheet_names else None
        
        if target_sheet:
            sheet_info = sheets[target_sheet]
            df = sheet_info["df"]
            columns = sheet_info["columns"]
            preview = sheet_info["preview"]
            
            code = generate_query_code(columns, preview, user_query, target_sheet)
            print("\n--- Generated Code ---")
            print(code)
            
            exec_result = execute_snippet(df, code)
            
            if exec_result["ok"]:
                result = exec_result["result"]
                print("\n--- Query Result ---")
                if isinstance(result, pd.DataFrame):
                    print(result.to_string(index=False))
                else:
                    print(json.dumps(result, indent=2, default=str))
                results["query_result"] = result
            else:
                print(f"❌ Error: {exec_result.get('error')}")
                results["errors"].append(exec_result.get("error"))
    
    # Save updated file if there were updates or new sheets
    if results["sheets_processed"] or results["new_sheets"]:
        updated_sheets = {name: sheets[name]["df"] for name in sheet_names}
        try:
            new_path = save_updated_excel(path, updated_sheets, results["new_sheets"])
            results["saved_path"] = new_path
        except Exception as e:
            error_msg = f"Failed to save file: {str(e)}"
            results["errors"].append(error_msg)
            print(f"❌ {error_msg}")
    
    return results

# ===============================
# Interactive Loop Tool
# ===============================
@tool(
    name="interactive_loop",
    description="Interactive loop to handle Excel operations across multiple sheets.",
    show_result=True
)
def interactive_loop(path: str, user_query: str) -> str:
    """Main entry point for the Excel agent."""
    print("\n" + "="*60)
    print("=== Multi-Sheet Excel Agent with Gemini ===")
    print("="*60)
    
    path = path.strip()
    if not path:
        return "❌ No path provided; exiting."
    
    print(f"✅ Processing file: {path}")
    
    user_query = user_query.strip()
    if user_query.lower() in {"exit", "quit"}:
        return "👋 Goodbye."
    
    try:
        results = process_query_all_sheets(path, user_query)
        
        summary = f"\n\n{'='*60}\n"
        summary += "📊 SUMMARY\n"
        summary += f"{'='*60}\n"
        summary += f"Intent: {results['intent']}\n"
        
        if results.get("sheets_processed"):
            summary += f"Sheets Updated: {', '.join(results['sheets_processed'])}\n"
        
        if results.get("new_sheets"):
            summary += f"New Sheets Created: {', '.join(results['new_sheets'].keys())}\n"
            for sheet_name, df in results['new_sheets'].items():
                summary += f"  - {sheet_name}: {len(df)} rows, {len(df.columns)} columns\n"
        
        if results.get("errors"):
            summary += f"\n⚠️  Errors: {len(results['errors'])}\n"
            for err in results['errors']:
                summary += f"  - {err}\n"
        
        if results.get("saved_path"):
            summary += f"\n💾 File saved: {results['saved_path']}\n"
        else:
            summary += f"\n⚠️  No changes made to file\n"
        
        print(summary)
        return summary
        
    except Exception as e:
        error_msg = f"❌ Fatal error: {str(e)}\n{traceback.format_exc()}"
        print(error_msg)
        return error_msg

# ===============================
# Main Excel Agent Setup
# ===============================
if __name__ == "__main__":
    # Create the agent
    excel_agent = Agent(
        model=Gemini(id="gemini-2.0-flash", api_key=os.getenv("GEMINI_API_KEY")),
        tools=[interactive_loop],
        instructions="""You are an advanced Excel agent that can:
        1. Work with multi-sheet Excel files
        2. Create new sheets with user-specified data
        3. Update existing sheets based on user queries
        4. Handle field dependencies automatically (e.g., if user increments employee count, update related calculations)
        5. Process queries across all sheets or specific sheets
        
        Always use the interactive_loop tool to handle user requests.""",
        markdown=True,
    )
    
    # Example usage
    excel_agent.print_response(
        "This is my excel path /Users/jayanth/Desktop/SheetShift/excel/Imperial_auto_faridabad_Solution file.xlsx "
        "I want to create a new sheet with highlighter, stapler, ruler, each with quantity 69, "
        "price 5, and discount 50%, and save the file.",
        stream=True
    )