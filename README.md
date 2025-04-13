# Autonomous Paint Agent

An intelligent agent that uses Google's Gemini 1.5 Flash model to control Microsoft Paint for visualizing ASCII exponential calculations.

## 📁 Project Structure
```
├── .env                    # Environment variables
├── .gitignore             # Git ignore rules
├── .python-version        # Python version specification
├── README.md             # Project documentation
├── autonomous_paint_agent.py  # Main agent implementation
├── paint_tools.py        # Paint control tools
├── pyproject.toml        # Project dependencies and metadata
└── uv.lock              # UV dependency lock file
```

## Overview

This project demonstrates an autonomous agent that:
1. Calculates ASCII exponential sums
2. Visualizes results in Microsoft Paint
3. Saves the visualization
4. Emails the result to specified recipients

## 📊 Prompt Evaluation

Using our Prompt Evaluation Assistant (Gemini 1.5), our system prompt scores:

```json
{
  "explicit_reasoning": true,
  "structured_output": true,
  "tool_separation": true,
  "conversation_loop": true,
  "instructional_framing": true,
  "internal_self_checks": true,
  "reasoning_type_awareness": true,
  "fallbacks": true,
  "overall_clarity": "Excellent structure with comprehensive self-checks and error handling"
}
```

## 🤖 Model & Architecture

- **LLM**: Google Gemini 1.5 Flash
- **Framework**: MCP (Model Context Protocol)
- **Key Components**: 
  - Autonomous execution
  - Tool-based architecture
  - Structured prompt engineering

## 🎯 Key Features

### 1. Explicit Reasoning Instructions
```python
## 🧠 Reasoning & Execution Instructions

1. **Think step-by-step.** Before calling any function, explain what you're doing and why.
2. **Label each reasoning type** using `REASONING: <type>`, e.g., `REASONING: arithmetic`
3. **Verify intermediate results** before moving on.
```

### 2. Structured Output Format
```python
FUNCTION_CALL: {"name": "function_name", "args": {"param1": "value1"}}
FINAL_ANSWER: <brief summary of what was done>
```

### 3. Tool Separation
The agent has access to distinct tools for:
- ASCII calculations (`ascii_exp_sum`)
- Paint manipulation (`open_paint`, `draw_rectangle`, `add_text_in_paint`)
- File operations (`save_paint_file`)
- Communication (`send_email_with_attachment`)

### 4. Internal Self-Checks
```python
# Example from our code
if hasattr(result, 'content'):
    if isinstance(result.content, list):
        result_text = result.content[0].text
    else:
        result_text = str(result.content)
```

### 5. Reasoning Type Awareness
```python
REASONING: arithmetic  
Calculated the sum of ASCII characters raised to increasing powers. Result verified.

REASONING: planning
Determining next visualization step based on canvas dimensions.
```

### 6. Error Handling & Fallbacks
```python
try:
    response = await generate_with_timeout(client, prompt)
except asyncio.TimeoutError:
    print("LLM generation timed out!")
    raise
except Exception as e:
    print(f"Error in LLM generation: {e}")
    raise
```

## 🎨 Visualization Specifications

- Canvas Size: 1600x808 pixels
- Rectangle Position: (600, 354) to (1000, 454)
- Text Position: Centered at (800, 404)

## 🚀 Getting Started

1. Create a `.env` file with:
   ```
   GEMINI_API_KEY=your_api_key
   GMAIL_ADDRESS=your_email
   ```

2. Install uv (if not already installed):
   ```bash
   pip install uv
   ```

3. Create and activate virtual environment:
   ```bash
   uv venv
   # On Windows:
   .venv/Scripts/activate
   # On Unix/MacOS:
   source .venv/bin/activate
   ```

4. Initialize project with uv:
   ```bash
   # Initialize new pyproject.toml and uv.lock
   uv init

   # Replace the generated pyproject.toml content with:
   [project]
   name = "Paint & Email Agent"
   version = "0.1.0"
   description = "Prompt Improvements, JSON Parsing and more!"
   requires-python = ">=3.10"
   dependencies = [
       "google-genai>=1.10.0",
       "mcp>=1.6.0",
       "python-dotenv>=1.1.0",
       "pywin32>=310",
       "pywinauto>=0.6.9",
       "psutil"
   ]
   ```

5. Install dependencies using uv sync:
   ```bash
   # Sync dependencies and update lock file
   uv sync
   ```

6. Run the agent:
   ```bash
   python autonomous_paint_agent.py
   ```

## 🔄 Conversation Loop Support

The agent maintains state through:
- Iteration tracking
- Response history
- Tool result caching

## ⚠️ Error Handling

The agent implements multiple layers of error handling:
1. Tool execution timeouts
2. LLM response validation
3. Resource cleanup
4. Graceful fallbacks

## 🔍 Code Quality Features

1. Comprehensive error handling
2. Clear separation of concerns
3. Structured prompt engineering
4. Self-documenting code
5. Resource management 
