
import os
from flask import Flask, request, jsonify, render_template
from dotenv import load_dotenv
from google import genai

# Load environment variables
load_dotenv()

app = Flask(__name__)

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

if not GEMINI_API_KEY:
    print("FATAL ERROR: GEMINI_API_KEY not found. Please set it in your .env file.")
    exit()

try:
    # Create Gemini client
    client = genai.Client(api_key=GEMINI_API_KEY)

    print("Gemini Client Initialized Successfully.")

except Exception as e:
    print(f"Error initializing Gemini Client: {e}")
    exit()


@app.route('/', methods=['POST', ])
def chat():
    # Handle GET requests
    if request.method == 'GET':
        return jsonify({"error": "This endpoint only accepts POST requests with a 'prompt' field"}), 405
    
    # Handle POST requests - try JSON first, then form data
    if request.is_json:
        data = request.get_json()
        user_prompt = data.get('prompt') if data else None
    else:
        # Try form data
        user_prompt = request.form.get('prompt')
        if not user_prompt:
            # Try to parse JSON manually if Content-Type wasn't set correctly
            try:
                data = request.get_json(force=True)
                user_prompt = data.get('prompt') if data else None
            except:
                user_prompt = None

    if not user_prompt:
        return jsonify({"error": "No prompt provided. Please send a POST request with a 'prompt' field in JSON or form data."}), 400

    try:
        # Force the model to produce EXACTLY one paragraph
        system_instruction = (
            "Write a clear, concise, well-structured paragraph (5–7 sentences) "
            "based on the prompt below. Do not produce multiple paragraphs, "
            "bullet points, headings, or lists. One paragraph only."
        )

        response = client.models.generate_text(
            model="gemini-2.0-flash",
            prompt=f"{system_instruction}\n\nUser prompt: {user_prompt}"
        )

        return jsonify({
            "response": response.text.strip()
        })

    except Exception as e:
        print(f"Gemini API Error: {e}")
        return jsonify({
            "error": "An error occurred while communicating with the AI model."
        }), 500



if __name__ == '__main__':
    app.run(debug=True)


