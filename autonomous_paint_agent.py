#!/usr/bin/env python
"""
Paint Agent - Uses LLM to control Microsoft Paint via MCP with autonomous execution order
"""
import json
import os
import sys
import asyncio
from dotenv import load_dotenv
from mcp import ClientSession, StdioServerParameters, types
from mcp.client.stdio import stdio_client
import traceback
from google import genai
import atexit
import signal

# Try to load .env from multiple possible locations
possible_env_paths = [
    ".env",  # Current directory
    "Assignment/.env",  # If run from Session4
]

env_loaded = False
for env_path in possible_env_paths:
    if os.path.exists(env_path):
        print(f"Found .env file at: {env_path}")
        load_dotenv(env_path)
        env_loaded = True
        break

if not env_loaded:
    print("Warning: Could not find .env file in any expected location")

# Access your API key and initialize Gemini client
api_key = os.getenv("GEMINI_API_KEY")
if not api_key:
    print("Error: GEMINI_API_KEY not found in environment variables")
    print("Please create a .env file with GEMINI_API_KEY=your_api_key")
    sys.exit(1)

client = genai.Client(api_key=api_key)

# Setup tracking for iterations
max_iterations = 10  # Increased to allow for more flexible execution
iteration = 0
iteration_response = []

async def generate_with_timeout(client, prompt, timeout=60):
    """Generate content with a timeout"""
    print("Starting LLM generation...")
    try:
        # Convert the synchronous generate_content call to run in a thread
        loop = asyncio.get_event_loop()
        response = await asyncio.wait_for(
            loop.run_in_executor(
                None, 
                lambda: client.models.generate_content(
                    model="gemini-1.5-flash",
                    contents=prompt
                )
            ),
            timeout=timeout
        )
        print("LLM generation completed")
        return response
    except asyncio.TimeoutError:
        print("LLM generation timed out!")
        raise
    except Exception as e:
        print(f"Error in LLM generation: {e}")
        raise

async def format_tools_for_prompt(tools):
    """Format tools information for the prompt"""
    if not tools:
        return "No tools available."
    
    tools_description = []
    for i, tool in enumerate(tools):
        try:
            # Get tool properties
            params = tool.inputSchema
            desc = getattr(tool, 'description', 'No description available')
            name = getattr(tool, 'name', f'tool_{i}')
            
            # Format the input schema in a more readable way
            if 'properties' in params:
                param_details = []
                for param_name, param_info in params['properties'].items():
                    param_type = param_info.get('type', 'unknown')
                    param_details.append(f"{param_name}: {param_type}")
                params_str = ', '.join(param_details)
            else:
                params_str = 'no parameters'

            tool_desc = f"{i+1}. {name}({params_str}) - {desc}"
            tools_description.append(tool_desc)
            print(f"Added description for tool: {tool_desc}")
        except Exception as e:
            print(f"Error processing tool {i}: {e}")
            tools_description.append(f"{i+1}. Error processing tool")
    
    return "\n".join(tools_description)

def cleanup_resources():
    """Clean up any remaining resources."""
    try:
        # Force flush any buffered output
        sys.stdout.flush()
        sys.stderr.flush()
        
        # Clean up any subprocesses if needed
        import psutil
        current_process = psutil.Process()
        children = current_process.children(recursive=True)
        for child in children:
            try:
                child.terminate()
            except Exception as e:
                print(f"Error terminating child process: {e}")
    except Exception as e:
        print(f"Cleanup error: {e}")

# Register the cleanup function
atexit.register(cleanup_resources)

async def main():
    """Main function to run the Paint agent with Gemini."""
    global iteration, iteration_response
    server_process = None
    
    print("Starting Paint Agent execution...")
    
    try:
        # Find the paint_tools.py script path
        possible_paths = [
            "Session5/Assignment/paint_tools.py",
            "paint_tools.py",
            "./paint_tools.py",
            "Assignment/paint_tools.py",  # When running from Session4
            os.path.join(os.getcwd(), "Session5/Assignment/paint_tools.py"),
            os.path.join(os.getcwd(), "paint_tools.py"),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "paint_tools.py"),  # Same dir as script
        ]
        
        path = None
        for p in possible_paths:
            if os.path.exists(p):
                path = p
                break
        
        if not path:
            print("Error: Could not find paint_tools.py")
            sys.exit(1)
        
        print(f"Found paint_tools.py at: {path}")
        
        # Create parameters for the server connection
        print("Starting MCP server...")
        server_params = StdioServerParameters(
            command="python",
            args=[path]
        )

        # Connect to MCP server
        print("Establishing connection to MCP server...")
        try:
            # Set a timeout for the server connection
            connection_timeout = 10  # seconds
            
            # Create a client connection with proper cleanup
            async with stdio_client(server_params) as (read, write):
                print("Connection established, creating session...")
                async with ClientSession(read, write) as session:
                    print("Session created, initializing...")
                    await session.initialize()
                    
                    # Get available tools
                    print("Requesting tool list...")
                    tools_result = await session.list_tools()
                    tools = tools_result.tools
                    print(f"Successfully retrieved {len(tools)} tools")

                    # Create system prompt with available tools
                    print("Creating system prompt...")
                    tools_description = await format_tools_for_prompt(tools)

                    # Call the list_tools method to get the tools
                    tools_result = await session.list_tools()

                    # Extract tool names from the tools_result
                    tool_names = [tool.name for tool in tools_result.tools]  # Use dot notation to access the name attribute

                    # Print the tool names for debugging
                    print("Available tools:", tool_names)
                    
                    # Create a prompt that will guide the LLM to control Paint and send emails
                    system_prompt = f"""You are an agent that can control Microsoft Paint and send emails.

---

## 🧰 Available Tools:
{tools_description}

---

## 🎯 Your Task

You must:

1. **Calculate the ASCII exponential sum for a given input string**  
2. **Visualize the result** in Microsoft Paint:  
   - Draw a rectangle centered at (800, 404) on a 1600*808 canvas.  
   - Rectangle coordinates: (600, 354) to (1000, 454)
   - Add the resulting number as text at (800, 404).
3. **Save the visualization** as an image file.  
4. **Email the image** to a specified recipient.

---

## 🧠 Reasoning & Execution Instructions

1. **Think step-by-step.** Before calling any function, explain what you're doing and why.
2. **Label each reasoning type** using `REASONING: <type>`, e.g., `REASONING: arithmetic`, `REASONING: planning`, `REASONING: verification`.
3. **Verify intermediate results** (e.g., validate calculations or output format) before moving on.
4. **Separate reasoning and tool-use clearly.** Always reason first, then call the function.
5. **If something goes wrong or you're uncertain**, pause, explain the issue, and try a fallback strategy.
6. **Sanity check your own outputs**—do they look valid, plausible, and match expectations?

---

## 🗂 Output Format

IMPORTANT: Respond with EXACTLY ONE line in this format, with NO markdown code blocks or other formatting
Do not add any datatypes in the function call:

FUNCTION_CALL: {{"name": "function_name", "args": {{"param1": "value1", "param2": "value2", "param3": "value3",....}}}}

When all steps are complete, respond with:

FINAL_ANSWER: <brief summary of what was done>

Do not include any extra commentary, explanation, or markdown in your responses.

---

## 📐 Visual Layout Requirements

- When drawing a rectangle:
  - It must surround the text centered at 800, 404
  - Use a rectangle like: 600, 354, 1000, 454 — providing ample space for long numbers
- When adding text:
  - Place it **exactly** at 800, 404
- When saving the file:
  - Use a consistent and descriptive filename, like `ascii_result_<input_string>.png`.
- When emailing:
  - **Use the exact same filename** used in the save step.
  - Refer to earlier steps to confirm the filename.

---

## 🧪 Example

Example response (note: no markdown, just plain text):
FUNCTION_CALL: {{"name": "ascii_exp_sum", "args": {{"input_string": "Hello World"}} }}

REASONING: arithmetic  
Calculated the sum of ASCII characters raised to increasing powers. Result verified.

FUNCTION_CALL: {{"name": "draw_rectangle", "args": {{"x1": "476", "y1": "274", "x2": "676", "y2": "374"}} }}

---

## 🔁 Multi-Turn Support

If this task spans multiple turns, remember your previous function calls and results. Use that context to continue logically. Always double-check filenames and prior outputs before proceeding.

---"""

                    # Get input string from the user
                    user_input = input("\nEnter a string to calculate its ASCII exponential sum: ")
                    if not user_input:
                        user_input = "Hello World"  # Default if user doesn't enter anything
                        print(f"Using default input: '{user_input}'")
                        
                    # Get recipient email from the user
                    recipient_email = input("\nEnter recipient email address for sending the visualization: ")
                    if not recipient_email:
                        recipient_email = os.getenv("GMAIL_ADDRESS", "")  # Default to own email if in .env
                        if recipient_email:
                            print(f"Using default recipient: {recipient_email}")
                        else:
                            print("No recipient email provided. Will skip email step.")

                    # Initial query
                    query = f"""Calculate the ASCII exponential sum for "{user_input}", visualize it in Paint, and email it to {recipient_email}."""

                    # Main interaction loop
                    while iteration < max_iterations:
                        print(f"\n=== Iteration {iteration + 1} ===")
                        
                        # Generate response
                        print("Generating LLM response...")
                        prompt = f"{system_prompt}\n\nQuery: {query}"
                        
                        try:
                            response = await generate_with_timeout(client, prompt)
                            response_text = response.text.strip()
                            print(f"LLM Response: {response_text}")
                            
                            # Split the response into lines and process each FUNCTION_CALL
                            lines = response_text.split('\n')
                            for line in lines:
                                line = line.strip()
                                if not line:
                                    continue
                                    
                                # Skip reasoning and other non-function lines
                                if not line.startswith("FUNCTION_CALL:") and not line.startswith("FINAL_ANSWER:"):
                                    continue
                                
                                if line.startswith("FUNCTION_CALL:"):
                                    # Parse the function call
                                    print(f"Line: {line}")
                                    _, function_info = line.split(":", 1)

                                   
                                    try:
                                        function_info_json = json.loads(function_info)
                                    except json.JSONDecodeError as e:
                                        print(f"JSON decoding error: {e}")
                                        print(f"Function info that caused the error: {function_info}")
                                    func_name, params = function_info_json['name'], function_info_json['args']
                                    
                                    print(f"Executing function: {func_name} with params: {params}")
                                    
                                    
                                    if func_name in tool_names:
                                        print(f"Calling tool {func_name} with arguments: {params}")
                                        try:
                                            # Call the tool and wait for result
                                            result = await session.call_tool(func_name, arguments=params)
                                            
                                            # Process the result
                                            if hasattr(result, 'content'):
                                                if isinstance(result.content, list):
                                                    result_text = result.content[0].text
                                                else:
                                                    result_text = str(result.content)
                                            else:
                                                result_text = str(result)
                                            
                                            # Add the result to the conversation history
                                            iteration_response.append(f"In iteration {iteration + 1}, {func_name} returned: {result_text}")
                                            print(f"Tool result: {result_text}")
                                            
                                            # Add delay between function calls
                                            await asyncio.sleep(2)
                                            
                                        except Exception as e:
                                            print(f"Error executing {func_name}: {e}")
                                            iteration_response.append(f"In iteration {iteration + 1}, {func_name} failed: {str(e)}")
                                
                                elif line.startswith("FINAL_ANSWER:"):
                                    print("\n=== Agent Execution Complete ===")
                                    print(line)
                                    return
                            
                            # Update the query with all results from this iteration
                            query = f"""Previous steps and results:
{chr(10).join(iteration_response)}

What should be the next step?"""
                            
                            print("Query: ", query)
                            
                        except Exception as e:
                            print(f"Error in iteration {iteration + 1}: {e}")
                            break
                            
                        iteration += 1
                        
                    if iteration >= max_iterations:
                        print("\n=== Maximum iterations reached ===")
                        print("The agent has reached the maximum number of iterations.")
                        
        except Exception as e:
            print(f"Error in MCP server connection: {e}")
            traceback.print_exc()
            return
    except Exception as e:
        print(f"Error in main execution: {e}")
        traceback.print_exc()
    finally:
        # Explicitly flush stdout/stderr to prevent pipe errors
        sys.stdout.flush()
        sys.stderr.flush()
        await asyncio.sleep(0.5)  # Give a moment for pipes to clear

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nExecution interrupted by user")
    except Exception as e:
        print(f"Fatal error: {e}")
        traceback.print_exc()
        sys.exit(1)
    finally:
        # Ensure all resources are properly cleaned up
        sys.stdout.flush()
        sys.stderr.flush()
        
        # Clean up any subprocesses
        try:
            import psutil
            current_process = psutil.Process()
            children = current_process.children(recursive=True)
            for child in children:
                try:
                    print(f"Terminating child process: {child.pid}")
                    child.terminate()
                except:
                    pass
        except Exception as e:
            print(f"Cleanup error: {e}") 