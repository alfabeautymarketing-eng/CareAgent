#!/usr/bin/env python3
"""
Mini-NotebookLM Tool
--------------------
This script mimics the functionality of Google NotebookLM by loading specific
context files (documentation, code) and allowing the user to ask questions
based on that context using the Gemini API.

Usage:
    python scripts/mini_notebook.py "What are the rules for beads?"
    python scripts/mini_notebook.py --interactive
"""

import sys
import os
import argparse
from pathlib import Path
from typing import List

# Add project root to sys.path to allow imports from src
project_root = Path(__file__).parent.parent
sys.path.append(str(project_root))

from src.services.gemini_client import get_gemini_client
from src.utils.logger import logger

# Default files to include in the "Notebook"
DEFAULT_CONTEXT_FILES = [
    "AI_RULES.md",
    "AGENTS.md",
    "README.md",
    "pyproject.toml",
]

def load_context(files: List[str]) -> str:
    """Load content from specified files."""
    context_parts = []
    
    for path_arg in files:
        # Support glob patterns if the shell didn't expand them
        # or if the user provided them in a way that requires expansion
        path_obj = project_root / path_arg
        
        # If it's an exact match (file or existing path), use it.
        # Otherwise, try to treat it as a glob pattern from root.
        matches = []
        if path_obj.exists():
            matches = [path_obj]
        else:
            # Try globbing from project root
            # Note: rglob for recursive if "**" is used, or just glob
            try:
                # We use glob on project_root
                matches = list(project_root.glob(path_arg))
            except Exception as e:
                logger.warning(f"Error globbing {path_arg}: {e}")
        
        if not matches:
             logger.warning(f"File or pattern not found: {path_arg}")
             continue

        for file_path in matches:
            try:
                if file_path.is_file():
                    # Check file size to avoid overloading context
                    if file_path.stat().st_size > 100_000: # Skip > 100KB
                        logger.warning(f"Skipping large file: {file_path.name}")
                        continue
                        
                    content = file_path.read_text(encoding="utf-8", errors="replace")
                    rel_path = file_path.relative_to(project_root)
                    context_parts.append(f"=== FILE: {rel_path} ===\n{content}\n")
            except Exception as e:
                logger.error(f"Error reading {file_path.name}: {e}")
            
    return "\n".join(context_parts)

def build_prompt(query: str, context: str) -> str:
    """Construct the prompt for Gemini."""
    return f"""You are an intelligent assistant helping a developer understand their project.
You have access to the following project documentation and files (Context).

CONTEXT:
{context}

USER QUESTION:
{query}

INSTRUCTIONS:
1. Answer the question STRICTLY based on the provided Context.
2. If the answer is not in the context, say so, but try to infer from related info if possible (marking it as inference).
3. Quote specific files or rules when relevant.
4. Be concise and practical.
"""

def ask_notebook(query: str, files: List[str] = None):
    """Run the query against the context."""
    if files is None:
        files = DEFAULT_CONTEXT_FILES
        
    print(f"📚 Loading context from {len(files)} files...")
    context = load_context(files)
    
    if not context:
        print("❌ No context loaded. Check file paths.")
        return

    print("🤖 Asking Gemini...")
    client = get_gemini_client()
    
    try:
        prompt = build_prompt(query, context)
        # Using a slightly higher temperature for more natural explanation, 
        # but still low for accuracy.
        response = client.call_api(
            prompt=prompt,
            temperature=0.2,
            max_output_tokens=4096 
        )
        
        print("\n" + "="*80)
        print(f"GURUS ANSWER (based on {', '.join(files)}):")
        print("="*80 + "\n")
        print(response)
        print("\n" + "="*80)
        
    except Exception as e:
        print(f"\n❌ Error calling Gemini API: {e}")

def main():
    parser = argparse.ArgumentParser(description="Mini-NotebookLM for AgentCare")
    parser.add_argument("query", nargs="?", help="The question to ask")
    parser.add_argument("--interactive", "-i", action="store_true", help="Run in interactive mode")
    parser.add_argument("--files", "-f", nargs="+", help="Specific files to use as context", default=None)
    
    args = parser.parse_args()
    
    files = args.files if args.files else DEFAULT_CONTEXT_FILES
    
    if args.interactive:
        print("Welcome to AgentCare Mini-NotebookLM! (Type 'exit' to quit)")
        while True:
            try:
                query = input("\n📝 Question: ")
                if query.lower() in ('exit', 'quit', 'q'):
                    break
                if not query.strip():
                    continue
                ask_notebook(query, files)
            except KeyboardInterrupt:
                break
    elif args.query:
        ask_notebook(args.query, files)
    else:
        parser.print_help()

if __name__ == "__main__":
    main()
