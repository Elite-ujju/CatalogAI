from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import google.generativeai as genai
import pandas as pd
import json
import io
import os
from dotenv import load_dotenv
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

load_dotenv()

app = Flask(__name__)
CORS(app)

GEMINI_MODELS = [
    "gemini-flash-latest",
    "gemini-flash-lite-latest",
    "gemini-2.5-flash",
    "gemini-2.5-flash-lite",
    "gemini-2.5-pro",
    "gemini-pro-latest",
]

def get_gemini_response(prompt, api_key):
    genai.configure(api_key=api_key)
    for model_name in GEMINI_MODELS:
        try:
            model = genai.GenerativeModel(model_name)
            response = model.generate_content(prompt)
            return response.text
        except Exception as e:
            continue
    raise Exception("All Gemini models failed. Check your API key.")

def analyze_product_prompt(product_data):
    prompt = f"""
You are an expert industrial product data analyst. Analyze this product and return a detailed JSON response.

Product Data Provided:
{json.dumps(product_data, indent=2)}

Instructions:
1. Identify the product type (Motor, Pump, Valve, Sensor, Transformer, Solar Panel, etc.)
2. List all provided fields and their values
3. Identify missing important fields for this product type
4. Try to fill missing fields using industry knowledge
5. Detect any conflicts or inconsistencies in the data
6. Cross-verify mathematical relationships
7. Calculate a readiness score

Return ONLY a valid JSON object with this exact structure:
{{
  "product_type": "string - identified product type",
  "product_name": "string - product name or description",
  "provided_fields": {{
    "field_name": "value"
  }},
  "enriched_fields": {{
    "field_name": {{
      "value": "enriched value",
      "confidence": 85,
      "source": "AI Knowledge Base / Industry Standard / Calculated",
      "human_check": false
    }}
  }},
  "missing_fields": [
    {{
      "field": "field name",
      "reason": "why this field is important",
      "suggested_value": "if possible",
      "confidence": 0
    }}
  ],
  "conflicts": [
    {{
      "field": "field name",
      "issue": "description of conflict",
      "severity": "HIGH / MEDIUM / LOW",
      "recommendation": "what to do"
    }}
  ],
  "cross_verification": [
    {{
      "check": "check name",
      "formula": "formula used",
      "status": "PASS / FAIL / CANNOT_VERIFY",
      "details": "explanation"
    }}
  ],
  "readiness_score": {{
    "total": 78,
    "grade": "B",
    "completeness": 75,
    "accuracy": 80,
    "verification": 70,
    "confidence": 85
  }},
  "summary": "2-3 sentence summary of the product and its data quality"
}}

Important:
- Be honest about confidence levels
- If confidence is below 60, set human_check to true
- Detect real mathematical conflicts
- Grade scale: A+ (95-100), A (85-94), B (70-84), C (55-69), D (40-54), F (below 40)
- Return ONLY the JSON, no extra text
"""
    return prompt

def parse_ai_response(response_text):
    try:
        cleaned = response_text.strip()
        if cleaned.startswith("```json"):
            cleaned = cleaned[7:]
        if cleaned.startswith("```"):
            cleaned = cleaned[3:]
        if cleaned.endswith("```"):
            cleaned = cleaned[:-3]
        cleaned = cleaned.strip()
        return json.loads(cleaned)
    except:
        return None

@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({"status": "ok", "message": "CatalogAI Server Running"})

@app.route('/api/verify-key', methods=['POST'])
def verify_key():
    data = request.json
    api_key = data.get('api_key', '')
    if not api_key:
        return jsonify({"success": False, "message": "API key is empty"})
    try:
        genai.configure(api_key=api_key)
        for model_name in GEMINI_MODELS:
            try:
                model = genai.GenerativeModel(model_name)
                response = model.generate_content("Say OK")
                return jsonify({"success": True, "message": "API key verified successfully"})
            except:
                continue
        return jsonify({"success": False, "message": "All models failed. Check your API key."})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)})

@app.route('/api/analyze-single', methods=['POST'])
def analyze_single():
    data = request.json
    api_key = data.get('api_key', '')
    product_data = data.get('product_data', {})
    if not api_key:
        return jsonify({"success": False, "message": "API key required"})
    if not product_data:
        return jsonify({"success": False, "message": "Product data required"})
    try:
        prompt = analyze_product_prompt(product_data)
        response_text = get_gemini_response(prompt, api_key)
        result = parse_ai_response(response_text)
        if result:
            return jsonify({"success": True, "data": result})
        else:
            return jsonify({"success": False, "message": "Could not parse AI response. Try again."})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)})

@app.route('/api/download-template', methods=['GET'])
def download_template():
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Product Data"
        headers = [
            "Product Name", "Product Type", "Brand", "Model Number",
            "Power (W)", "Voltage (V)", "Current (A)", "Frequency (Hz)",
            "RPM", "Efficiency (%)", "Power Factor", "IP Rating",
            "Flow Rate (m3/hr)", "Head (m)", "Pressure Rating (bar)",
            "Temperature Range", "Weight (kg)", "Dimensions",
            "Material", "Certifications", "Country of Origin", "Notes"
        ]
        header_fill = PatternFill(start_color="1a1a2e", end_color="1a1a2e", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.column_dimensions[get_column_letter(col_num)].width = 18
        ws.row_dimensions[1].height = 30
        sample_data = [
            ["Industrial Motor 5HP", "Motor", "Siemens", "1LA7090-2AA10",
             3730, 415, 8.5, 50, 2900, 91.5, 0.85, "IP55",
             "", "", "", "-20 to 40 C", 28, "250x200x180mm",
             "Cast Iron", "CE, ISI", "Germany", "3 Phase Induction Motor"],
            ["Centrifugal Pump", "Pump", "Kirloskar", "STAR-1",
             1500, 230, 7.2, 50, 2900, 72, 0.8, "IP44",
             18, 15, "", "0 to 60 C", 12, "300x150x200mm",
             "Cast Iron", "ISI", "India", ""],
            ["Gate Valve DN50", "Valve", "L&T", "GV-50",
             "", "", "", "",
             "", "", "", "",
             "", "", 16, "-10 to 120 C", 3.5, "DN50",
             "Stainless Steel", "API 600", "India", "Full bore gate valve"],
            ["Pressure Sensor", "Sensor", "Honeywell", "PX2AN1XX150PAAAX",
             "", 24, 0.02, "",
             "", "", "", "IP67",
             "", "", 150, "-40 to 85 C", 0.15, "M20x1.5",
             "Stainless Steel", "CE, ATEX", "USA", "0-150 PSI range"],
            ["Solar Panel 400W", "Solar Panel", "Waaree", "WS-400",
             400, 48.2, 10.12, "",
             "", 20.5, "", "IP68",
             "", "", "", "-40 to 85 C", 21, "1755x1038x35mm",
             "Anodized Aluminum", "IEC 61215", "India", "Monocrystalline"],
        ]
        row_colors = ["F8F9FA", "FFFFFF"]
        for row_idx, row_data in enumerate(sample_data, 2):
            fill_color = row_colors[row_idx % 2]
            row_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.fill = row_fill
                cell.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[row_idx].height = 22
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='CatalogAI_Template.xlsx'
        )
    except Exception as e:
        return jsonify({"success": False, "message": str(e)})

@app.route('/api/process-excel', methods=['POST'])
def process_excel():
    api_key = request.form.get('api_key', '')
    if not api_key:
        return jsonify({"success": False, "message": "API key required"})
    if 'file' not in request.files:
        return jsonify({"success": False, "message": "No file uploaded"})
    file = request.files['file']
    if file.filename == '':
        return jsonify({"success": False, "message": "No file selected"})
    try:
        df = pd.read_excel(file)
        if df.empty:
            return jsonify({"success": False, "message": "Excel file is empty"})

        # Collect all products first
        all_products = []
        for idx, row in df.iterrows():
            product_data = {}
            for col in df.columns:
                val = row[col]
                if pd.notna(val) and str(val).strip() != '':
                    product_data[col] = str(val)
            if product_data:
                all_products.append({"row_index": idx + 1, "data": product_data})

        if not all_products:
            return jsonify({"success": False, "message": "No valid data found in Excel"})

        # Send ALL products in ONE AI call (FAST!)
        batch_prompt = f"""
You are an expert industrial product data analyst. Analyze ALL {len(all_products)} products below and return a JSON array with detailed analysis for each.

Products to analyze:
{json.dumps([p['data'] for p in all_products], indent=2)}

For EACH product, return an object with this exact structure:
{{
  "product_type": "identified type",
  "product_name": "product name",
  "provided_fields": {{"field": "value"}},
  "enriched_fields": {{
    "field_name": {{
      "value": "enriched value",
      "confidence": 85,
      "source": "AI Knowledge Base / Industry Standard / Calculated",
      "human_check": false
    }}
  }},
  "missing_fields": [
    {{"field": "name", "reason": "why important", "suggested_value": "if possible", "confidence": 0}}
  ],
  "conflicts": [
    {{"field": "name", "issue": "description", "severity": "HIGH/MEDIUM/LOW", "recommendation": "action"}}
  ],
  "cross_verification": [
    {{"check": "name", "formula": "formula", "status": "PASS/FAIL/CANNOT_VERIFY", "details": "explanation"}}
  ],
  "readiness_score": {{
    "total": 78,
    "grade": "B",
    "completeness": 75,
    "accuracy": 80,
    "verification": 70,
    "confidence": 85
  }},
  "summary": "2-3 sentence summary"
}}

Return ONLY a valid JSON array like: [{{...product1...}}, {{...product2...}}, ...]
No markdown, no extra text, just the JSON array.
Grade scale: A+ (95-100), A (85-94), B (70-84), C (55-69), D (40-54), F (below 40).
If confidence below 60, set human_check to true.
"""
        try:
            response_text = get_gemini_response(batch_prompt, api_key)
            cleaned = response_text.strip()
            if cleaned.startswith("```json"):
                cleaned = cleaned[7:]
            if cleaned.startswith("```"):
                cleaned = cleaned[3:]
            if cleaned.endswith("```"):
                cleaned = cleaned[:-3]
            cleaned = cleaned.strip()

            batch_results = json.loads(cleaned)

            results = []
            for i, ai_result in enumerate(batch_results):
                if i < len(all_products):
                    ai_result['row_index'] = all_products[i]['row_index']
                    ai_result['original_data'] = all_products[i]['data']
                    results.append(ai_result)

            return jsonify({"success": True, "results": results, "total": len(results)})

        except Exception as batch_err:
            results = []
            for product in all_products:
                try:
                    prompt = analyze_product_prompt(product['data'])
                    response_text = get_gemini_response(prompt, api_key)
                    result = parse_ai_response(response_text)
                    if result:
                        result['row_index'] = product['row_index']
                        result['original_data'] = product['data']
                        results.append(result)
                    else:
                        results.append({
                            "row_index": product['row_index'],
                            "original_data": product['data'],
                            "error": "Could not parse AI response"
                        })
                except Exception as e:
                    results.append({
                        "row_index": product['row_index'],
                        "original_data": product['data'],
                        "error": str(e)
                    })
            return jsonify({"success": True, "results": results, "total": len(results)})

    except Exception as e:
        return jsonify({"success": False, "message": str(e)})

@app.route('/api/download-results', methods=['POST'])
def download_results():
    try:
        data = request.json
        results = data.get('results', [])
        if not results:
            return jsonify({"success": False, "message": "No results to download"})

        wb = openpyxl.Workbook()

        # Sheet 1: Enriched Products
        ws1 = wb.active
        ws1.title = "Enriched Products"
        headers1 = ["Row", "Product Name", "Product Type", "Brand", "Grade",
                    "Total Score", "Completeness", "Accuracy", "Verification",
                    "Confidence", "Summary"]
        header_fill = PatternFill(start_color="1a1a2e", end_color="1a1a2e", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for col, h in enumerate(headers1, 1):
            cell = ws1.cell(row=1, column=col, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
            ws1.column_dimensions[get_column_letter(col)].width = 20
        for row_idx, r in enumerate(results, 2):
            if 'error' not in r:
                score = r.get('readiness_score', {})
                ws1.cell(row=row_idx, column=1, value=r.get('row_index', ''))
                ws1.cell(row=row_idx, column=2, value=r.get('product_name', ''))
                ws1.cell(row=row_idx, column=3, value=r.get('product_type', ''))
                ws1.cell(row=row_idx, column=4, value=r.get('original_data', {}).get('Brand', ''))
                ws1.cell(row=row_idx, column=5, value=score.get('grade', ''))
                ws1.cell(row=row_idx, column=6, value=score.get('total', ''))
                ws1.cell(row=row_idx, column=7, value=score.get('completeness', ''))
                ws1.cell(row=row_idx, column=8, value=score.get('accuracy', ''))
                ws1.cell(row=row_idx, column=9, value=score.get('verification', ''))
                ws1.cell(row=row_idx, column=10, value=score.get('confidence', ''))
                ws1.cell(row=row_idx, column=11, value=r.get('summary', ''))

        # Sheet 2: Conflicts
        ws2 = wb.create_sheet("Conflicts Report")
        headers2 = ["Row", "Product", "Field", "Issue", "Severity", "Recommendation"]
        for col, h in enumerate(headers2, 1):
            cell = ws2.cell(row=1, column=col, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
            ws2.column_dimensions[get_column_letter(col)].width = 25
        conflict_row = 2
        for r in results:
            if 'error' not in r:
                for conflict in r.get('conflicts', []):
                    ws2.cell(row=conflict_row, column=1, value=r.get('row_index', ''))
                    ws2.cell(row=conflict_row, column=2, value=r.get('product_name', ''))
                    ws2.cell(row=conflict_row, column=3, value=conflict.get('field', ''))
                    ws2.cell(row=conflict_row, column=4, value=conflict.get('issue', ''))
                    ws2.cell(row=conflict_row, column=5, value=conflict.get('severity', ''))
                    ws2.cell(row=conflict_row, column=6, value=conflict.get('recommendation', ''))
                    conflict_row += 1

        # Sheet 3: Enrichment Details
        ws3 = wb.create_sheet("Enrichment Details")
        headers3 = ["Row", "Product", "Field", "Enriched Value", "Confidence %", "Source", "Human Check"]
        for col, h in enumerate(headers3, 1):
            cell = ws3.cell(row=1, column=col, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
            ws3.column_dimensions[get_column_letter(col)].width = 25
        enrich_row = 2
        for r in results:
            if 'error' not in r:
                for field, details in r.get('enriched_fields', {}).items():
                    ws3.cell(row=enrich_row, column=1, value=r.get('row_index', ''))
                    ws3.cell(row=enrich_row, column=2, value=r.get('product_name', ''))
                    ws3.cell(row=enrich_row, column=3, value=field)
                    ws3.cell(row=enrich_row, column=4, value=details.get('value', ''))
                    ws3.cell(row=enrich_row, column=5, value=details.get('confidence', ''))
                    ws3.cell(row=enrich_row, column=6, value=details.get('source', ''))
                    ws3.cell(row=enrich_row, column=7, value="Yes" if details.get('human_check') else "No")
                    enrich_row += 1

        # Sheet 4: Summary
        ws4 = wb.create_sheet("Summary")
        ws4.column_dimensions['A'].width = 35
        ws4.column_dimensions['B'].width = 20
        title_font = Font(bold=True, size=14, color="1a1a2e")
        ws4.cell(row=1, column=1, value="CatalogAI Processing Summary").font = title_font
        ws4.cell(row=3, column=1, value="Total Products Processed").font = Font(bold=True)
        ws4.cell(row=3, column=2, value=len(results))
        successful = [r for r in results if 'error' not in r]
        ws4.cell(row=4, column=1, value="Successfully Analyzed").font = Font(bold=True)
        ws4.cell(row=4, column=2, value=len(successful))
        if successful:
            scores = [r.get('readiness_score', {}).get('total', 0) for r in successful]
            ws4.cell(row=5, column=1, value="Average Readiness Score").font = Font(bold=True)
            ws4.cell(row=5, column=2, value=round(sum(scores) / len(scores), 1))
            total_conflicts = sum(len(r.get('conflicts', [])) for r in successful)
            ws4.cell(row=6, column=1, value="Total Conflicts Found").font = Font(bold=True)
            ws4.cell(row=6, column=2, value=total_conflicts)
            total_enriched = sum(len(r.get('enriched_fields', {})) for r in successful)
            ws4.cell(row=7, column=1, value="Total Fields Enriched").font = Font(bold=True)
            ws4.cell(row=7, column=2, value=total_enriched)

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='CatalogAI_Results.xlsx'
        )
    except Exception as e:
        return jsonify({"success": False, "message": str(e)})

@app.route('/api/chat', methods=['POST'])
def chat():
    data = request.json
    api_key = data.get('api_key', '')
    user_message = data.get('message', '')
    chat_history = data.get('history', [])

    if not api_key:
        return jsonify({"success": False, "message": "API key required"})
    if not user_message:
        return jsonify({"success": False, "message": "Message required"})

    history_text = ""
    for msg in chat_history[-6:]:
        role = "User" if msg.get('role') == 'user' else "Assistant"
        history_text += f"{role}: {msg.get('content', '')}\n"

    chat_prompt = f"""
You are CatalogAI Assistant — a friendly, knowledgeable AI helper embedded in the CatalogAI website.

YOUR PERSONALITY:
- Friendly, warm, and helpful (like a smart friend who happens to be an industrial expert)
- Answer ANY question the user asks — don't refuse general topics
- Be conversational, not robotic
- Keep answers SHORT (2-6 sentences typically)
- Use simple English

YOUR PRIMARY EXPERTISE (be extra detailed here):

1. INDUSTRIAL PRODUCTS & B2B:
   - Motors, Pumps, Valves, Sensors, Transformers, Bearings, VFDs, etc.
   - Specifications, IP ratings, IE classes, standards (CE, ISI, BIS, ATEX, IEC, ISO)
   - How to choose, compare, or troubleshoot industrial products
   - B2B procurement, MOQ, lead times, supplier evaluation
   - Indian industrial market (HSN codes, IS standards, GST)
   - Manufacturing processes and industrial safety

2. SOLUTION & SUGGESTION EXPERT:
   - When user asks about ANY industrial product, also automatically tell them:
     - Common faults/issues that product typically has
     - How to solve or prevent those faults
     - Better alternatives available in market
     - Maintenance tips for longer product life
     - Safety warnings if any
     - Cost/energy optimization suggestions
   - Example: User asks "Tell me about 5HP motor"
     Answer includes: specs + "Common issues: overheating due to overload, bearing failure. Solutions: Install thermal overload relay, regular lubrication every 6 months. Better alternative: IE3 efficiency motor saves 15% energy."

3. CATALOGAI WEBSITE HELP:
   - How to upload Excel files
   - What is Readiness Score, Grade, Confidence, Human Check
   - How AI analysis works (cleaning, enriching, conflict detection)
   - How to interpret results and download reports
   - Excel template structure
   - Troubleshooting

BUT ALSO — ANSWER OTHER QUESTIONS TOO:
- General knowledge questions (history, science, technology, math, etc.) — answer them helpfully
- Business questions (marketing, startup, growth, industries) — give a good answer
- Casual conversation (greetings, how are you, jokes) — be friendly
- Coding/tech questions — help if you can
- Educational questions — teach clearly
- Career advice — give useful guidance
- General news/trends — share what you know, but mention if info might be outdated
- Weather / real-time data — politely say you can't check live data, but suggest checking Google/apps

HOW TO ANSWER:

For INDUSTRIAL/WEBSITE questions:
→ Give detailed, expert answer with faults, solutions, and alternatives

For GENERAL questions:
→ Answer helpfully like ChatGPT would
→ At the end, if natural, mention how CatalogAI could help (optional, don't force it)

For REAL-TIME questions (weather, live news, current stock prices):
→ Say: "I don't have live data access, but you can check [Google/app name]. However, I can share general information about [topic]..."

For OFF-TOPIC or WEIRD questions:
→ Still answer politely and helpfully
→ Don't be rigid or refuse

RESPONSE RULES:
- Never say "I can only help with X"
- Always try to be helpful
- Keep tone friendly and natural
- Use dashes (-) for lists, not asterisks
- No markdown symbols like ** or ##
- Answer in the SAME language user uses (English → English, Hindi → Hindi, Hinglish → Hinglish)

Previous conversation:
{history_text}

User's new question: {user_message}

Provide a helpful, friendly response:
"""

    try:
        response_text = get_gemini_response(chat_prompt, api_key)
        return jsonify({"success": True, "reply": response_text.strip()})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)})

if __name__ == '__main__':
    print("CatalogAI Server Starting...")
    print("Open: http://localhost:5000")
    app.run(debug=True, port=5000)