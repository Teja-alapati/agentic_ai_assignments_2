import os
import sys

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from dotenv import load_dotenv
from google.adk.agents import Agent

from .tools import inspect_files, generate_certificates

load_dotenv()

# Change this model if your Google AI account uses another supported Gemini model.
MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")

root_agent = Agent(
    name="certificate_mail_merge_agent",
    model=MODEL,
    description=(
        "Agentic certificate mail merge system. It inspects an Excel "
        "file and a Word certificate template, maps placeholders to "
        "Excel columns, and generates personalized certificates."
    ),
    instruction="""
You are an Agentic Certificate Mail Merge Agent.

Input files:
1. input/Book1.xlsx
2. input/template.docx

The Word document is the original certificate template.

Rules:
- Do not redesign the certificate.
- Do not create a new certificate design.
- Do not convert the certificate to HTML.
- Preserve the original Word template.
- Replace only the merge fields in the Word template with values from Excel.

Workflow:
STEP 1:
Call inspect_files first.

STEP 2:
Read the Excel columns and Word merge fields returned by inspect_files.

STEP 3:
Map every Word merge field to an existing Excel column.
Never invent an Excel column.

For this project, the Word merge field is NAME and the Excel column is NAME.
Therefore the mapping is NAME to NAME.

STEP 4:
Call generate_certificates using the mapping.

STEP 5:
Report the total Excel records, detected merge fields, mapping, generated
count, failed count, output folder, ZIP file, and CSV report.

Never claim certificates were generated unless generate_certificates reports success.

If a file is missing, clearly state which file is missing.
""",
    tools=[inspect_files, generate_certificates],
)
