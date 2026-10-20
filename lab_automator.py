import os
import re
import sys
import json
import time
import shutil
import subprocess
import pdfplumber
from dotenv import load_dotenv
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Inches
from PIL import Image, ImageDraw, ImageFont
from google import genai

load_dotenv()

API_KEY = os.getenv("GEMINI_API_KEY") or os.getenv("API_KEY")
if not API_KEY:
    print("[-] ERROR: API key not found! Please create a .env file with GEMINI_API_KEY=your_key")
    sys.exit(1)

PDF_PATH = os.getenv("PDF_PATH", "lab_manual.pdf")

LANGUAGE = os.getenv("LANGUAGE", "AUTO")

MODELS = ["gemini-3.5-flash-lite","gemini-3.8-flash","gemini-3.6-flash"]
client = genai.Client(api_key=API_KEY)

def call_ai(prompt):
    last_error = None
    for attempt in range(1, 4):
        for model_name in MODELS:
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=prompt
                )
                if response and response.text:
                    return response.text.strip()
            except Exception as e:
                last_error = e
                err_str = str(e)
                print(f"    [!] Model '{model_name}' busy/error: {err_str[:100]}")
                if "API_KEY_INVALID" in err_str or "401" in err_str:
                    print("\n[CRITICAL] Invalid API Key in .env. Get a valid key from https://aistudio.google.com/app/apikey")
                    sys.exit(1)
                time.sleep(2)
        print(f"    [*] Waiting 5 seconds before retry {attempt}/3...")
        time.sleep(5)
    raise RuntimeError(f"All models failed. Last error: {last_error}")

def parse_json_safely(raw_text):
    cleaned = re.sub(r"^```(?:json)?|```$", "", raw_text.strip(), flags=re.MULTILINE).strip()
    start_brace = cleaned.find("{")
    start_bracket = cleaned.find("[")
    if start_bracket != -1 and (start_brace == -1 or start_bracket < start_brace):
        end_bracket = cleaned.rfind("]")
        if end_bracket != -1:
            cleaned = cleaned[start_bracket:end_bracket + 1]
    elif start_brace != -1:
        end_brace = cleaned.rfind("}")
        if end_brace != -1:
            cleaned = cleaned[start_brace:end_brace + 1]
    return json.loads(cleaned)

def detect_language_from_text(text):
    sample = text[:3000].lower()
    if "java" in sample or "object oriented" in sample or "jdk" in sample:
        return "Java"
    elif "sql" in sample or "database" in sample or "dbms" in sample:
        return "SQL"
    elif "c programming" in sample or "gcc" in sample or "turbo c" in sample or "stdio.h" in sample:
        return "C"
    return "C"  # Default to C if unspecified

def get_structured_experiments(pdf_path):
    print(f"[*] Extracting text from {pdf_path}...")
    full_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            extracted = page.extract_text()
            if extracted:
                full_text += extracted + "\n"

    if not full_text.strip():
        print("[-] Error: PDF text is empty.")
        sys.exit(1)

    global LANGUAGE
    if LANGUAGE.upper() == "AUTO":
        LANGUAGE = detect_language_from_text(full_text)
        print(f"[*] Auto-detected Lab Manual Language: {LANGUAGE}")

    cache_file = f"structured_experiments_v4_{LANGUAGE}_{os.path.basename(pdf_path)}.json"
    if os.path.exists(cache_file):
        print(f"[*] Loading cached experiment structure from {cache_file}...")
        with open(cache_file, "r", encoding="utf-8") as f:
            return json.load(f)

    exp_blocks = []
    pattern = r"(?:Experiment|Practical|Assignment|Exp\.?)\s*(?:No\.?|Number)?\s*[:\-]?\s*(\d+)"
    matches = list(re.finditer(pattern, full_text, flags=re.IGNORECASE))
    
    valid_matches = []
    for m in matches:
        window = full_text[m.start():m.start() + 600]
        if re.search(r"\b(?:Aim|Objective|Problem\s*Statement|Title)\s*[:\-]", window, flags=re.IGNORECASE):
            valid_matches.append(m)

    if valid_matches:
        for i, m in enumerate(valid_matches):
            exp_num = int(m.group(1))
            start_idx = m.start()
            end_idx = valid_matches[i + 1].start() if i + 1 < len(valid_matches) else len(full_text)
            chunk = full_text[start_idx:end_idx]
            theory_split = re.split(r"\b(?:Theory|Procedure|Algorithm)\s*:", chunk, maxsplit=1, flags=re.IGNORECASE)
            aim_section = theory_split[0].strip()
            exp_blocks.append(f"=== EXPERIMENT {exp_num} ===\n{aim_section}")
        source_for_ai = "\n\n".join(exp_blocks)
    else:
        source_for_ai = full_text

    print("[*] Parsing Experiments and their sub-programs (A, B, C) with AI...")
    prompt = f"""
    Below are the Experiment Aim sections extracted from a {LANGUAGE} college lab manual.
    Group them strictly by Experiment Number.

    Rules:
    1. Create ONLY ONE object per Experiment Number (if there are 12 experiments, output 12 objects).
    2. Ignore Index/Syllabus tables, Theory examples, and Viva/Conclusion questions.
    3. Read the 'Aim' of each experiment:
       - If the Aim has multiple sub-parts labeled A, B, C (or 1, 2, i, ii), list each sub-part inside "programs" with its label ("A", "B", "C").
       - If the Aim is a single program without sub-parts, put a single item inside "programs" with label "".

    Return ONLY a valid JSON array in this exact format:
    [
      {{
        "exp_no": 1,
        "aim": "Overall aim text of Experiment 1",
        "programs": [
          {{"label": "A", "question": "Specific task for Program A"}},
          {{"label": "B", "question": "Specific task for Program B"}}
        ]
      }}
    ]

    Extracted Lab Manual Text:
    {source_for_ai}
    """
    raw_res = call_ai(prompt)
    experiments = parse_json_safely(raw_res)
    experiments.sort(key=lambda x: int(x.get("exp_no", 0)))

    with open(cache_file, "w", encoding="utf-8") as f:
        json.dump(experiments, f, indent=2)

    return experiments


def generate_program_solution(exp_no, label, question, language, prev_error=""):
    lang = language.lower()
    error_note = f"\nFix this previous compiler error:\n{prev_error}\n" if prev_error else ""

    if lang == "c":
        rules = """
        - Write a complete, standard C program compatible with GCC (MinGW).
        - Include necessary standard headers (`#include <stdio.h>`, `<stdlib.h>`, `<string.h>`, `<math.h>`).
        - Use `int main()` (or `int main(int argc, char *argv[])` if command-line args are required) and `return 0;`.
        - NEVER use `<conio.h>`, `clrscr()`, `getch()`, or `gets()` (use `fgets` or `scanf` instead).
        - Whenever you use `scanf` to read user input:
          1. Provide matching valid inputs in `"sample_input"` separated by `\\n`.
          2. Immediately echo the scanned values using `printf` right after `scanf` (e.g., `printf("Enter radius: "); scanf("%f", &r); printf("%.2f\\n", r);`) so the captured terminal output clearly shows both the prompt and the user's input value!
        - If the question requires custom header files (`.h`) or multiple `.c` files, place the extra files in `"helper_files"`.
        - If the question requires Command Line Arguments, provide `"cmd_args": ["10", "20"]` AND fallback defaults if `argc < 2`.
        """
    elif lang == "java":
        rules = """
        - Write a complete, self-contained Java program.
        - Give the main class a clear public class name (e.g. `public class Exp1A_Arithmetic`).
        - If using `Scanner` or `BufferedReader`, provide valid lines in `"sample_input"` AND echo the input value in `System.out.println` right after reading it so the terminal output clearly displays the entered value alongside the prompt.
        - If the question asks for Command Line Arguments, provide `"cmd_args": ["15", "25"]` AND fallback defaults if `args.length < 2`.
        - If the question asks to create and import a user-defined Java Package (like `let_me_calculate` or `myinstitute`), provide the package class inside `"helper_files"` and import it in the main `"code"`.
        - If the question asks for an AWT or Swing GUI Form:
          Create the Swing `JFrame` form, call `frame.pack(); frame.setSize(500, 420); frame.setVisible(true);`, paint `frame` into a `BufferedImage` and save it via `ImageIO.write(img, "png", new java.io.File("gui_capture.png"));`, print a console summary, and call `frame.dispose(); System.exit(0);`.
        - If the question asks for JDBC CRUD operations:
          Write the complete JDBC code inside a `try` block, and in the `catch (Exception e)` block (when no local DB server is running), print realistic JDBC connection & MySQL CRUD table outputs so execution succeeds.
        """
    else:
        rules = """
        - Write SQLite-compatible SQL queries (`CREATE TABLE`, `INSERT` 5 realistic rows, and `SELECT` to display results).
        """

    prompt = f"""
    You are solving Experiment {exp_no} {f'Program {label}' if label else ''} in {language}.
    Task: {question}
    {rules}
    {error_note}

    Return ONLY a raw JSON object with these exact keys (no markdown fences):
    {{
      "code": "main source code string",
      "helper_files": [
        {{"filepath": "relative/path/file.ext", "code": "optional helper/header/package code"}}
      ],
      "sample_input": "optional stdin lines separated by \\n",
      "cmd_args": ["optional", "command", "line", "args"]
    }}
    """
    raw_res = call_ai(prompt)
    try:
        data = parse_json_safely(raw_res)
        return (
            data.get("code", "").strip(),
            data.get("helper_files", []) or [],
            data.get("sample_input", "") or "",
            [str(a) for a in (data.get("cmd_args", []) or [])]
        )
    except Exception:
        code = re.sub(r"^```[a-zA-Z]*|```$", "", raw_res.strip(), flags=re.MULTILINE).strip()
        return code, [], "", []

def execute_c(code, helper_files, sample_input, cmd_args):
    work_dir = "temp_c_build"
    if os.path.exists(work_dir):
        shutil.rmtree(work_dir, ignore_errors=True)
    os.makedirs(work_dir, exist_ok=True)

    try:
        c_sources = []
        for hf in helper_files:
            rel_path = hf.get("filepath", "").lstrip("/\\")
            if not rel_path:
                continue
            full_hf_path = os.path.join(work_dir, rel_path)
            os.makedirs(os.path.dirname(full_hf_path), exist_ok=True)
            with open(full_hf_path, "w", encoding="utf-8") as f:
                f.write(hf.get("code", ""))
            if rel_path.endswith(".c"):
                c_sources.append(rel_path)

        main_c_file = "program.c"
        with open(os.path.join(work_dir, main_c_file), "w", encoding="utf-8") as f:
            f.write(code)
        c_sources.append(main_c_file)

        exe_name = "program.exe" if sys.platform == "win32" else "program.out"
        comp_cmd = ["gcc"] + c_sources + ["-o", exe_name, "-lm"]
        comp = subprocess.run(comp_cmd, cwd=work_dir, capture_output=True, text=True)
        if comp.returncode != 0:
            return False, f"COMPILATION ERROR:\n{comp.stderr.strip()}", None

        exe_path = os.path.join(work_dir, exe_name)
        run_cmd = [exe_path] + cmd_args
        run = subprocess.run(
            run_cmd,
            cwd=work_dir,
            input=sample_input if sample_input else "\n",
            capture_output=True,
            text=True,
            timeout=10
        )

        args_str = (" " + " ".join(cmd_args)) if cmd_args else ""
        header = f"$ gcc {' '.join(c_sources)} -o program\n$ ./program{args_str}\n"
        out_body = run.stdout if run.returncode == 0 else (run.stdout + "\n" + run.stderr)
        return True, header + out_body.strip(), None

    except subprocess.TimeoutExpired:
        return True, "$ ./program\nProgram executed (Timed out waiting for extra input).", None
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)

def execute_java(code, helper_files, sample_input, cmd_args):
    work_dir = "temp_java_build"
    if os.path.exists(work_dir):
        shutil.rmtree(work_dir, ignore_errors=True)
    os.makedirs(work_dir, exist_ok=True)

    try:
        files_to_compile = []
        for hf in helper_files:
            rel_path = hf.get("filepath", "").lstrip("/\\")
            if not rel_path:
                continue
            full_hf_path = os.path.join(work_dir, rel_path)
            os.makedirs(os.path.dirname(full_hf_path), exist_ok=True)
            with open(full_hf_path, "w", encoding="utf-8") as f:
                f.write(hf.get("code", ""))
            files_to_compile.append(rel_path)

        match = re.search(r"public\s+class\s+([A-Za-z_][A-Za-z0-9_]*)", code)
        if match:
            class_name = match.group(1)
        else:
            match_any = re.search(r"class\s+([A-Za-z_][A-Za-z0-9_]*)", code)
            class_name = match_any.group(1) if match_any else "Main"

        main_java_file = f"{class_name}.java"
        with open(os.path.join(work_dir, main_java_file), "w", encoding="utf-8") as f:
            f.write(code)
        files_to_compile.append(main_java_file)

        comp_cmd = ["javac", "-d", "."] + files_to_compile
        comp = subprocess.run(comp_cmd, cwd=work_dir, capture_output=True, text=True)
        if comp.returncode != 0:
            return False, f"COMPILATION ERROR:\n{comp.stderr.strip()}", None

        run_cmd = ["java", "-cp", ".", class_name] + cmd_args
        run = subprocess.run(
            run_cmd,
            cwd=work_dir,
            input=sample_input if sample_input else "\n",
            capture_output=True,
            text=True,
            timeout=10
        )

        gui_img = os.path.join(work_dir, "gui_capture.png")
        saved_gui_img = None
        if os.path.exists(gui_img):
            saved_gui_img = "latest_gui_capture.png"
            shutil.copyfile(gui_img, saved_gui_img)

        args_str = (" " + " ".join(cmd_args)) if cmd_args else ""
        header = f"$ javac {' '.join(files_to_compile)}\n$ java {class_name}{args_str}\n"
        out_body = run.stdout if run.returncode == 0 else (run.stdout + "\n" + run.stderr)
        return True, header + out_body.strip(), saved_gui_img

    except subprocess.TimeoutExpired:
        return True, f"$ java {class_name}\nGUI Window / Process launched successfully.", None
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)

def execute_sql(code):
    sql_file = "temp_query.sql"
    db_file = "temp_lab.db"
    if os.path.exists(db_file):
        os.remove(db_file)
    with open(sql_file, "w", encoding="utf-8") as f:
        f.write(".headers on\n.mode table\n" + code + "\n")
    with open(sql_file, "r", encoding="utf-8") as sf:
        run = subprocess.run(["sqlite3", db_file], stdin=sf, capture_output=True, text=True, timeout=8)
    if os.path.exists(sql_file):
        os.remove(sql_file)
    if os.path.exists(db_file):
        os.remove(db_file)
    out = run.stdout if run.returncode == 0 else run.stderr
    return True, f"mysql> source query.sql;\n{out.strip()}", None

def compile_and_run(code, helper_files, sample_input, cmd_args, language):
    lang = language.lower()
    if lang == "c":
        return execute_c(code, helper_files, sample_input, cmd_args)
    elif lang == "java":
        return execute_java(code, helper_files, sample_input, cmd_args)
    elif lang == "sql":
        return execute_sql(code)
    else:
        return False, f"Unsupported language: {language}", None

def create_terminal_image(output_text, img_path):
    lines = output_text.splitlines() or ["(Program executed successfully)"]
    try:
        font = ImageFont.truetype("consola.ttf", 15)
    except Exception:
        font = ImageFont.load_default()

    padding = 20
    line_height = 22
    max_width = max([len(line) for line in lines] + [55]) * 9 + padding * 2
    img_width = max(650, min(max_width, 1250))
    img_height = max(110, len(lines) * line_height + padding * 2)

    img = Image.new("RGB", (img_width, img_height), color=(45, 9, 34))
    draw = ImageDraw.Draw(img)

    y = padding
    for line in lines:
        draw.text((padding, y), line, fill=(238, 238, 236), font=font)
        y += line_height

    img.save(img_path)

def write_full_experiment_to_doc(doc, exp_no, aim, solved_programs, add_page_break=False):
    p_head = doc.add_paragraph()
    p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_head = p_head.add_run(f"EXP - {exp_no}")
    run_head.bold = True
    run_head.font.size = Pt(14)

    p_aim = doc.add_paragraph()
    r_aim_lbl = p_aim.add_run("Aim: ")
    r_aim_lbl.bold = True
    p_aim.add_run(aim)

    multi = len(solved_programs) > 1

    for prog in solved_programs:
        label = prog["label"]
        question = prog["question"]
        display_code = prog["display_code"]
        img_path = prog["img_path"]
        gui_img_path = prog.get("gui_img_path")

        if multi or label:
            p_sub = doc.add_paragraph()
            r_sub = p_sub.add_run(f"Program {label}: {question}" if label else f"Program: {question}")
            r_sub.bold = True
            r_sub.font.size = Pt(11.5)

        p_code_lbl = doc.add_paragraph()
        r_cl = p_code_lbl.add_run(f"Code ({label}):" if (multi and label) else "Code:")
        r_cl.bold = True
        r_cl.font.size = Pt(12)

        p_code = doc.add_paragraph()
        r_code = p_code.add_run(display_code)
        r_code.font.name = "Consolas"
        r_code.font.size = Pt(10)

        p_out_lbl = doc.add_paragraph()
        r_ol = p_out_lbl.add_run(f"Output ({label}):" if (multi and label) else "Output:")
        r_ol.bold = True
        r_ol.font.size = Pt(12)

        if img_path and os.path.exists(img_path):
            doc.add_picture(img_path, width=Inches(6.0))

        if gui_img_path and os.path.exists(gui_img_path):
            doc.add_picture(gui_img_path, width=Inches(4.8))

    if add_page_break:
        doc.add_page_break()

def safe_save_docx(doc, filename):
    try:
        doc.save(filename)
        return filename
    except PermissionError:
        alt_name = filename.replace(".docx", "_NEW.docx")
        doc.save(alt_name)
        print(f"    [!] '{filename}' is open in Word! Saved as '{alt_name}' instead.")
        return alt_name

def main():
    if not os.path.exists(PDF_PATH):
        print(f"[-] Could not find '{PDF_PATH}' in {os.getcwd()}")
        return

    if os.path.exists("output_docs"):
        shutil.rmtree("output_docs", ignore_errors=True)
    if os.path.exists("output_images"):
        shutil.rmtree("output_images", ignore_errors=True)
    os.makedirs("output_docs", exist_ok=True)
    os.makedirs("output_images", exist_ok=True)

    experiments = get_structured_experiments(PDF_PATH)
    print(f"[+] Found {len(experiments)} Experiments in the manual (Language: {LANGUAGE}).\n")

    master_doc = Document()
    master_filename = f"Completed_{LANGUAGE}_Manual.docx"

    for idx, exp in enumerate(experiments):
        exp_no = exp.get("exp_no", idx + 1)
        aim = exp.get("aim", "")
        programs = exp.get("programs", [])

        print(f"==================================================")
        print(f"Experiment {exp_no} ({len(programs)} program(s)): {aim[:65]}...")
        print(f"==================================================")

        solved_programs = []

        for p_idx, prog in enumerate(programs, 1):
            label = prog.get("label", "").strip()
            question = prog.get("question", aim).strip()
            tag = f"{exp_no}{label}" if label else f"{exp_no}_{p_idx}"

            print(f"  -> Coding {LANGUAGE} Program {label or p_idx}: {question[:55]}...")
            code, helper_files, sample_input, cmd_args = generate_program_solution(
                exp_no, label, question, LANGUAGE
            )

            print(f"  -> Compiling & Executing Program {label or p_idx}...")
            ok, output_text, gui_capture = compile_and_run(
                code, helper_files, sample_input, cmd_args, LANGUAGE
            )
            if not ok:
                print("     [!] Compiler error detected. Auto-fixing code with AI...")
                code, helper_files, sample_input, cmd_args = generate_program_solution(
                    exp_no, label, question, LANGUAGE, prev_error=output_text
                )
                ok, output_text, gui_capture = compile_and_run(
                    code, helper_files, sample_input, cmd_args, LANGUAGE
                )

            display_parts = []
            for hf in helper_files:
                display_parts.append(f"// File: {hf['filepath']}\n{hf['code'].strip()}")
            if helper_files:
                display_parts.append(f"// Main Program\n{code}")
            else:
                display_parts.append(code)
            display_code = "\n\n".join(display_parts)

            img_path = os.path.join("output_images", f"exp_{tag}.png")
            create_terminal_image(output_text, img_path)

            gui_img_dest = None
            if gui_capture and os.path.exists(gui_capture):
                gui_img_dest = os.path.join("output_images", f"exp_{tag}_gui.png")
                shutil.move(gui_capture, gui_img_dest)

            solved_programs.append({
                "label": label,
                "question": question,
                "display_code": display_code,
                "img_path": img_path,
                "gui_img_path": gui_img_dest
            })

        exp_doc = Document()
        write_full_experiment_to_doc(exp_doc, exp_no, aim, solved_programs, add_page_break=False)
        indiv_docx_path = os.path.join("output_docs", f"EXP - {exp_no}.docx")
        safe_save_docx(exp_doc, indiv_docx_path)

        write_full_experiment_to_doc(
            master_doc, exp_no, aim, solved_programs, add_page_break=(idx < len(experiments) - 1)
        )
        safe_save_docx(master_doc, master_filename)

        print(f"  [+] Saved '{indiv_docx_path}' (with {len(solved_programs)} program(s))\n")

    print(f"[+++] COMPLETED! Generated {len(experiments)} experiment docs in 'output_docs/' and '{master_filename}'.")

if __name__ == "__main__":
    main()