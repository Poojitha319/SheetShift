from tools.exceltool import interactive_loop 
from agno.agent import Agent
from agno.models.google import Gemini
import os
import json

# Interactive loop to handle user input
# def interactive_loop():
#     print("=== Agno + Gemini Excel Agent ===")
    
#     # Check if GEMINI_API_KEY is set
#     if not os.environ.get("GEMINI_API_KEY"):
#         print("WARNING: GEMINI_API_KEY environment variable not set!")
#         print("Please set it with: export GEMINI_API_KEY='your-api-key'")
#         return
    
#     path = input("Enter path to Excel file: ").strip()
#     if not path:
#         print("No path provided; exiting.")
#         return
#     print(f"Loaded: {path}")
    
#     while True:
#         q = input("\nAsk a question about this Excel (type 'exit' to quit):\n> ").strip()
#         if not q:
#             continue
#         if q.lower() in {"exit", "quit"}:
#             print("Goodbye.")
#             break
        
#         print("\n🤖 Generating code with Gemini...")
#         out = answer_excel_question(path, q)
        
#         print("\n--- Generated Code ---")
#         print(out["code"])
#         print("\n--- Execution Result ---")
#         exec_out = out["execution"]
#         if exec_out.get("ok"):
#             result = exec_out["result"]
#             if isinstance(result, list) and len(result) > 0:
#                 print(f"✓ Found {len(result)} result(s):")
#             print(json.dumps(result, indent=2, default=str))
#         else:
#             print("❌ Error:", exec_out.get("error"))
#             if exec_out.get("traceback"):
#                 print(exec_out["traceback"])

# if __name__ == "__main__":
#     interactive_loop()



excel_agent = Agent(
    
    model=Gemini(id="gemini-2.0-flash",api_key = os.getenv("GEMINI_API_KEY")),
    tools=[interactive_loop],
    instructions="You are an Excel agent who can work with complex Excel files containing one or more sheets. You can use the interactive_loop tool to handle user input.",
    markdown=True,
)
excel_agent.print_response(" this is my excel path /Users/jayanth/Desktop/SheetShift/Invoice_20rows.xlsx i want to create a new sheet i want highlighter , stapler , ruler , of quantity 69 each and each price is 5 and discount is 50percent and save the file", stream=True)