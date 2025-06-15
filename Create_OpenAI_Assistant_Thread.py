import openai
import os

openai.api_key = os.getenv("OPENAI_API_KEY") or input("Enter your OpenAI API Key: ") or "hardcode_key_here"

# Create a thread
thread = openai.beta.threads.create()
thread_id = thread.id

print(f"Your thread ID is: {thread_id}")

with open("C:\\Scripts\\pingcastle_thread.txt", "w") as f:
    f.write(thread_id)
