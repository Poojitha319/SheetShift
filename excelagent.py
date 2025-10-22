from tools.exceltool import answer_excel_question
import os
import json

# Interactive loop to handle user input
def interactive_loop():
    print("=== Agno + Gemini Excel Agent ===")
    
    # Check if GEMINI_API_KEY is set
    if not os.environ.get("GEMINI_API_KEY"):
        print("WARNING: GEMINI_API_KEY environment variable not set!")
        print("Please set it with: export GEMINI_API_KEY='your-api-key'")
        return
    
    path = input("Enter path to Excel file: ").strip()
    if not path:
        print("No path provided; exiting.")
        return
    print(f"Loaded: {path}")
    
    while True:
        q = input("\nAsk a question about this Excel (type 'exit' to quit):\n> ").strip()
        if not q:
            continue
        if q.lower() in {"exit", "quit"}:
            print("Goodbye.")
            break
        
        print("\n🤖 Generating code with Gemini...")
        out = answer_excel_question(path, q)
        
        print("\n--- Generated Code ---")
        print(out["code"])
        print("\n--- Execution Result ---")
        exec_out = out["execution"]
        if exec_out.get("ok"):
            result = exec_out["result"]
            if isinstance(result, list) and len(result) > 0:
                print(f"✓ Found {len(result)} result(s):")
            print(json.dumps(result, indent=2, default=str))
        else:
            print("❌ Error:", exec_out.get("error"))
            if exec_out.get("traceback"):
                print(exec_out["traceback"])

if __name__ == "__main__":
    interactive_loop()