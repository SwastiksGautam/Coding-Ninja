import os
from fastapi import FastAPI
from pydantic import BaseModel
from typing import Dict, Any, List
from dotenv import load_dotenv
from openai import OpenAI
from fastapi.middleware.cors import CORSMiddleware

load_dotenv()

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

interview_state: Dict[str, Any] = {}

QUESTIONS = [
    "What is the difference between a formula and a function in Excel? Give an example of each.",
    "Explain the purpose of a Pivot Table and when you would use one.",
    "Describe how you would use the VLOOKUP or XLOOKUP function. What are their main limitations?",
    "How would you use a combination of IF and AND functions to create a conditional formula?",
    "What is data validation in Excel, and how does it improve data integrity?"
]

class ChatRequest(BaseModel):
    user_id: str = "test_user"
    message: str

def get_llm_response(messages: List[Dict[str, str]], final_prompt: bool = False):
    if final_prompt:
        system_prompt = (
            "You are a professional Excel interviewer. Review the entire conversation below. "
            "Based on the candidate's answers, provide constructive feedback on their performance. "
            "Finally, give them an estimated score out of 100. "
            "Example: 'Overall, your performance was excellent... Score: 90/100.'"
        )
    else:
        system_prompt = (
            "You are a professional Excel technical interviewer. Acknowledge the user's response concisely and then ask the next question."
        )

    full_conversation = [{"role": "system", "content": system_prompt}] + messages
    
    try:
        completion = client.chat.completions.create(
            model="gpt-4.1-nano-2025-04-14",
            messages=full_conversation,
            max_tokens=250,
            temperature=0.6
        )
        response_content = completion.choices[0].message.content.strip()
        return response_content
    except Exception as e:
        print(f"Error calling OpenAI API: {e}")
        return "Sorry, I'm having trouble connecting right now. Please try again later."

@app.post("/chat")
async def chat_endpoint(request_body: ChatRequest):
    user_id = request_body.user_id
    user_message = request_body.message
    
    if user_message.lower() == "start" and user_id not in interview_state:
        interview_state[user_id] = {
            "history": [],
            "phase": "conceptual",
            "turn": 0
        }
        intro_message = "Hello! I'm your AI Excel Interviewer. I'll ask you a few questions about Excel concepts first, then provide a practical task. Let's begin."
        interview_state[user_id]["history"].append({"role": "assistant", "content": intro_message})
        return {"response": intro_message, "status": "ongoing"}

    state = interview_state.get(user_id)
    if not state:
        return {"response": "Please click the start button to begin the interview.", "status": "not_started"}
    
    state["history"].append({"role": "user", "content": user_message})

    # This is the key change to introduce a conversational break
    if state["turn"] == 0:
        first_question = QUESTIONS[0]
        state["history"].append({"role": "assistant", "content": first_question})
        state["turn"] += 1
        return {"response": first_question, "status": "ongoing"}

    if state["turn"] <= len(QUESTIONS):
        llm_response = QUESTIONS[state["turn"]-1]
        state["history"].append({"role": "assistant", "content": llm_response})
        state["turn"] += 1
        return {"response": llm_response, "status": "ongoing"}

    elif state["turn"] == len(QUESTIONS) + 1:
        final_feedback = get_llm_response(state["history"], final_prompt=True)
        del interview_state[user_id]
        return {"response": final_feedback, "status": "complete"}
    
    # Fallback for unexpected turns
    return {"response": "I'm ready for the next question. Please continue.", "status": "ongoing"}

@app.get("/")
def read_root():
    return {"message": "Hello, this is the AI-Powered Excel Mock Interviewer Backend!"}

# You will add the other endpoints here later
# @app.get("/download_task")
# async def download_task():
#     pass

# @app.post("/upload_solution")
# async def upload_solution(file: UploadFile = File(...)):
#     pass