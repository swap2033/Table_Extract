import fitz  # PyMuPDF
import pandas as pd

def extract_text_from_pdf(pdf_path):
    # Open the PDF file
    document = fitz.open(pdf_path)
    all_text = []

    # Iterate through all the pages
    for page_num in range(len(document)):
        page = document.load_page(page_num)
        text = page.get_text("text")
        all_text.append(text)

    return all_text

def parse_tables(text_list):
    tables = []
    for text in text_list:
        lines = text.split('\n')
        data = [line.split() for line in lines if line.strip() != '']
        tables.append(pd.DataFrame(data))
    return tables

def pdf_to_excel(pdf_path, excel_path):
    # Extract text from the PDF
    text_list = extract_text_from_pdf(pdf_path)
    
    # Parse the extracted text into tables
    tables = parse_tables(text_list)
    
    # Create an Excel writer object
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        # Write each table to a separate sheet in the Excel file
        for i, table in enumerate(tables):
            table.to_excel(writer, sheet_name=f'Table_{i+1}', index=False)
            print(f"Table {i+1} written to sheet Table_{i+1}")
    
    print(f"Data has been written to {excel_path}")

# Example usage
pdf_path = 'path/to/your/file.pdf'
excel_path = 'path/to/save/your/file.xlsx'
pdf_to_excel(pdf_path, excel_path)


https://github.com/swap2033/MacroPython/

blob/main/Usage

Yes—you can effectively "freeze" your premium token/credit burn at zero while still actively using Copilot Chat and Agent mode by switching to GitHub's 0x (included) models or routing requests through BYOK (Bring Your Own Key) / local models.
1. Switch to 0x Included Models (Zero Quota Burn)
On any paid Copilot plan (Pro, Pro+, Business, Enterprise), GitHub designates specific base models with a 0x multiplier. Using them does not deduct from your monthly AI Credits or premium request pool:
 * Which models are 0x: Look at the multiplier tag next to the model name in the Copilot Chat model picker—models such as GPT-5 mini, Raptor mini, GPT-4.1, and GPT-4o are included at 0x.
 * How to use this to "pause" burn: Whenever you want to save your remaining credits for later in the month, click the model dropdown in Copilot Chat and select one of the 0x models instead of Auto, Claude Sonnet/Opus, or Gemini Pro. You can keep chatting, asking questions, and editing code without touching your remaining quota. (Note: 0x models are still subject to standard anti-abuse rate limits if you hammer them heavily.)
2. Use BYOK (Bring Your Own Key) or Local Models in VS Code
If you want to use frontier models (or local models) without touching your GitHub Copilot credit balance at all, VS Code's Copilot Chat lets you plug in external providers:
 * Local Models via Ollama / LM Studio (100% Free): Open the Copilot Chat model picker \rightarrow Manage Models... \rightarrow select Ollama (or OpenAI-compatible local endpoint). Any chat or edit sent to a local model uses zero GitHub tokens/credits.
 * Free-Tier External APIs (BYOK): In Manage Models..., you can add an API key from providers that offer generous free tiers (such as Google AI Studio / Gemini or OpenRouter free models). When that model is selected in the dropdown, requests go directly to your API key rather than your GitHub Copilot allowance.
3. Hybrid "Two-Model" Workflow (Best Way to Stretch Credits)
Instead of burning 1x–3x credits on every step of a task, split your workflow:
 * Draft & Boilerplate (0x model): Use GPT-5 mini or Raptor mini (0x) to scaffold files, write unit tests, explain unfamiliar code, and fix simple syntax/type errors.
 * Complex Debugging Only (1x+ model): Switch to a 1x model (like Claude Sonnet or Gemini Pro) only when the 0x model gets stuck on a multi-file architectural bug—then switch right back to 0x to freeze your meter again.

Because GitHub Copilot moved to usage-based billing (GitHub AI Credits) in June 2026, every token—input, cached input, and output—deducts directly from your monthly credit pool ($15 on Pro, $70 on Pro+, $200 on Max). Since fallback models were removed, freezing or near-freezing your burn while keeping the models running requires routing around your paid credit pool or leveraging micro-cost tiers:
1. Drop to Ultra-Low-Cost "Micro" Models ($0.10 / 1M Tokens)
While heavy frontier models like GPT-6 Astra ($10 input / $50 output per 1M tokens) or Claude Opus 5 ($5 / $25) can wipe out a Pro allowance in a single afternoon, micro-tier models inside Copilot barely move the meter:
 * Switch to GPT-6 Luna or Claude Haiku 5.5: Both cost $0.10 per 1M input tokens and $0.50 per 1M output tokens—literally 50x to 100x cheaper than flagship models. On a $15 Pro plan, you can run millions of tokens through Luna or Haiku 5.5 for pennies.
 * Avoid Long-Context Thresholds: Several models double in price once a single prompt crosses the standard context window into "long-context" pricing (for example, GPT-5.6 Luna jumps from $0.20 to $0.40 input, and Claude Haiku 5.5 jumps 5x from $0.10 to $0.50). Keeping chats short avoids triggering that surcharge.
2. Stay in "Cached Input" Territory (90%–95% Discount)
Under Copilot's token billing, cached input tokens cost 90% to 95% less than fresh input tokens:
 * Example: On GPT-6.1 Sol, fresh input costs $2.00 / 1M, while cached input is $0.10 / 1M (a 95% discount). On GPT-6 Luna, cached input drops to $0.01 / 1M.
 * How to exploit caching: Prompt caches live for roughly 5 minutes of inactivity and break whenever you change files at the top of your context or switch models mid-thread. Batch your follow-up questions on the same file within a 3–4 minute window using the same model so Copilot bills 90%+ of your context at the cached rate instead of re-reading everything from scratch.
3. Abuse Unlimited Inline Completions & Next Edit Suggestions
Unlike Chat, Agent mode, CLI, and PR Code Reviews—which all burn AI Credits—standard inline ghost-text completions and Next Edit Suggestions (NES) remain 100% unlimited and consume zero credits on all paid plans.
 * Comment-Driven Coding: Instead of opening Copilot Chat to ask "Write a function that parses this JWT and handles expiry," write a 1-line comment directly in your file (// Parse JWT, validate exp claim, return null if expired) and let Next Edit Suggestions + inline ghost-text write the block for free.
 * Chain Edits with Tab: When refactoring a variable or signature, make the first change manually and hit Tab repeatedly through Next Edit Suggestions rather than asking Agent mode to refactor the file.
4. Offload Chat & Agent Loops to Free External Providers (BYOK)
VS Code's Copilot Chat lets you connect external API keys (Chat: Manage Language Models) so you can keep using Copilot's UI, @workspace context, and Agent tools with zero GitHub credit burn:
 * Google AI Studio (Free Tier): Generate a free Gemini API key in Google AI Studio and plug it into VS Code's Copilot model manager. Rate-limited free requests let you run chat and edits without touching your GitHub balance.
 * OpenRouter Free Models / Groq / Mistral: Add an OpenAI-compatible endpoint in VS Code using free-tier routes (such as OpenRouter's :free models or Groq's free developer tier) to pause GitHub spend during heavy experimentation.
 * Local Models via Ollama / LM Studio: Run Qwen2.5-Coder, DeepSeek-R1, or Llama locally on your machine and select them inside Copilot Chat for completely offline, zero-credit coding.
5. Turn Off Dual-Meter Features (PR Code Review)
GitHub Copilot Pull Request Code Review now burns two meters simultaneously: GitHub AI Credits plus GitHub Actions compute minutes. Disable automatic Copilot reviews on draft PRs or every push in your repo settings, and trigger reviews manually only when a PR is ready to merge.
6. Set a Hard $0 Overage Budget Cap
Because fallback models no longer exist once your included credits hit zero, go to GitHub Settings \rightarrow Billing & Licensing \rightarrow Budgets and Alerts and confirm your AI Credits budget is capped at your plan's included amount ($0 overage). When you hit 100%, paid models pause automatically without charging your card while your unlimited inline completions and NES keep working.

 * 
