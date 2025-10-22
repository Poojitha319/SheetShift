import os
from textwrap import dedent
from agno.agent import Agent
from agno.models.google import Gemini
from agno.tools.file import FileTools



tools = [
    FileTools(
        enable_read_file=True,
        enable_save_file=True,
        enable_list_files=False
    )
]


os.environ["GEMINI_API_KEY"] = "your_gemini_api_key_here"
agent = Agent(
    model=Gemini(id="gemini-2.0-flash",api_key = os.getenv("GEMINI_API_KEY")),
    description=dedent("""
        You are a Universal Text Intelligence Agent that can analyze and transform any type of text data —
        including structured documents, configuration files, markdowns, logs, CSV-like data, JSON-like text,
        or unstructured paragraphs.

        Your main objective is to perform any text-based operation requested by the user while 
        **strictly preserving the original structure, indentation, layout, and data integrity**.

        Capabilities:
        1. Understand and manipulate any textual pattern (tables, code, logs, reports, configs, etc.).
        2. Accept complex natural-language instructions such as:
           - "Find all IP addresses and mask the last octet"
           - "Convert this changelog into markdown bullet format"
           - "Extract error logs and summarize by frequency"
           - "Correct grammar but keep spacing and indentation identical"
           - "Translate only comments to English"
           - "Reformat JSON-like text into valid JSON"
           - "Append summary of data insights at the end"
        3. Perform intelligent operations such as:
           - Pattern extraction and replacement
           - Reformatting while maintaining indentation and structure
           - Contextual text generation, summarization, or expansion
           - Content filtering, anonymization, or transformation
        4. For multi-section or structured data, apply changes consistently.
        5. Maintain every space, newline, and indentation unless explicitly told otherwise.

        Constraints:
        - Never destroy original structure or alignment unless user asks.
        - Do not reformat unless it's part of the task.
        - When uncertain, make a logical assumption and state it in the response.
        - Save the updated version with the same format as input.

        Output:
        - The modified file with all requested operations performed.
        - Optionally, include a summary of changes if requested by the user.
    """),
    tools=tools,
    retries=2,
)




input_path = "Ancient Indian astronomers built Yantras like the Samrat Yantra and Rama Yantra to observe celestial bodies and measure time with remarkable accuracy. These instruments, precisely calibrated for local coordinates in places like Jaipur and Ujjain, reflect India’s advanced astronomical knowledge. The challenge is to develop software that can generate their dimensions for any given latitude and longitude in India. The aim of this project is to design a web application that automates this process, preserving ancient scientific principles through modern computational modelling"               
output_path = "modified_output.txt"


user_prompt = """
1. Remove any duplicate lines.
2. Replace the word 'ERROR' with 'WARNING'.
3. At the end of the file, add a short summary of how many warnings were found.
4. Keep the line spacing and indentation identical.
"""



instruction = dedent(f"""
1. Read the file '{input_path}' using read_file().
2. Perform the following operations as per the user request: {user_prompt}
3. Preserve the original structure, indentation, and formatting exactly as it is.
4. Save the modified file to '{output_path}' using save_file().
5. If applicable, include a summary of what modifications were made.
""")


response = agent.run(instruction)

print("=== 🧾 AGENT RESPONSE ===")
print(response)





