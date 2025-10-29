from tools.exceltool import interactive_loop 
from agno.agent import Agent
from agno.models.google import Gemini
import os

excel_agent = Agent(
    
    model=Gemini(id="gemini-2.0-flash",api_key = os.getenv("GEMINI_API_KEY")),
    tools=[interactive_loop],
    instructions="You are an Excel agent who can work with complex Excel files containing one or more sheets. You can use the interactive_loop tool to handle user input.",
    markdown=True,
)
excel_agent.print_response(" this is my excel path  /Users/jayanth/Desktop/SheetShift/Imperial_auto_faridabad_Solution file.xlsx  i want to create a new sheet i want highlighter , stapler , ruler , of quantity 69 each and each price is 5 and discount is 50percent and save the file", stream=True)