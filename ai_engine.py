import json
import re
from excel_parser import DIM_ORDER

# ── System Prompt (same proven prompt from the Claude artifact) ──
SYSTEM_PROMPT = """You are a senior deal deliverability analyst at Accenture Security Practice. Generate PowerPoint content for Deal Deliverability Review.

RULES:
- Synthesize BOTH RFP and Excel assessment. Never repeat questions as bullets.
- Dimension bullets: 3 per dimension, each under 35 words.
- Comments & Actions: ONE short sentence per dimension.
- Key Justification: explain the overall RAG referencing specific dimensions.
- Positive Notes: exactly 4 short bullets.
- Assumptions: exactly 4, engagement-specific from RFP scope.
- Next Steps: exactly 2 with title, description, and owner role.
- Deal Overview: synthesize Excel overview + RFP into max 5 concise lines.

Respond ONLY valid JSON (no markdown, no backticks, no explanation):
{"opportunity_value":"High|Medium|Low","key_justification":"...","deal_overview":["line1","line2"],"positive_notes":["...","...","...","..."],"dimensions":[{"id":1,"name":"...","bullets":["...","...","..."],"comments":"..."},{"id":2,"name":"...","bullets":["..."],"comments":"..."},{"id":3,"name":"...","bullets":["..."],"comments":"..."},{"id":4,"name":"...","bullets":["..."],"comments":"..."},{"id":5,"name":"...","bullets":["..."],"comments":"..."}],"amber_summary":"...","red_summary":"...","assumptions":["...","...","...","..."],"next_steps":[{"title":"...","desc":"...","owner":"..."},{"title":"...","desc":"...","owner":"..."}]}"""


def _build_user_prompt(questions, overview, risks, rags, rfp_text):
    """Build the user prompt from parsed data."""
    
    # Assessment text
    assessment_lines = []
    for q in questions:
        line = f"Q{q['id']}[{q['dim']}]\"{q['question']}\"->{q['response']}|Team:{q.get('team', '-')}|RAG:{q.get('rag', '')}"
        if q.get('justification'):
            line += f"|J:{q['justification']}"
        if q.get('action'):
            line += f"|A:{q['action']}"
        assessment_lines.append(line)
    
    # Risks text
    risk_lines = [f"Risk{i+1}:{r['risk']}|Mit:{r.get('mit', '-')}" for i, r in enumerate(risks)]
    
    # Dimension RAGs
    dim_rag_str = ', '.join(f"{d}:{rags['dimRags'][d]}" for d in DIM_ORDER)
    
    prompt = f"""CALCULATED RAGs (use as given, do NOT recalculate):
Overall:{rags['overall']}|{dim_rag_str}
Rules applied: RED in dim=dim RED. AMBER no RED=AMBER. All GREEN=GREEN. RED dim=Overall RED.

DEAL OVERVIEW:{overview}
ASSESSMENT:
{chr(10).join(assessment_lines)}
RISKS:
{chr(10).join(risk_lines)}
RFP:
{rfp_text[:10000]}

Generate the PPT content JSON. Overall is {rags['overall']}. Key Justification MUST explain why it is {rags['overall']}."""
    
    return prompt


def _parse_json_response(text):
    """Extract and parse JSON from AI response, handling markdown fences."""
    # Remove markdown code fences if present
    text = re.sub(r'```json\s*', '', text)
    text = re.sub(r'```\s*', '', text)
    text = text.strip()
    
    # Find the JSON object
    start = text.find('{')
    end = text.rfind('}')
    if start == -1 or end == -1:
        raise ValueError("No JSON object found in AI response")
    
    json_str = text[start:end + 1]
    return json.loads(json_str)


def _call_gemini(api_key, system_prompt, user_prompt):
    """Call Google Gemini API."""
    import google.generativeai as genai
    
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(
        model_name='gemini-2.0-flash',
        system_instruction=system_prompt
    )
    
    response = model.generate_content(
        user_prompt,
        generation_config=genai.types.GenerationConfig(
            max_output_tokens=4000,
            temperature=0.3
        )
    )
    
    return response.text


def _call_groq(api_key, system_prompt, user_prompt):
    """Call Groq API (Llama 3)."""
    from groq import Groq
    
    client = Groq(api_key=api_key)
    response = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        max_tokens=4000,
        temperature=0.3
    )
    
    return response.choices[0].message.content


def _call_cohere(api_key, system_prompt, user_prompt):
    """Call Cohere API."""
    import cohere
    
    client = cohere.ClientV2(api_key=api_key)
    response = client.chat(
        model="command-r-plus",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        max_tokens=4000,
        temperature=0.3
    )
    
    return response.message.content[0].text


def generate_content(provider, api_key, questions, overview, risks, rags, rfp_text):
    """Generate AI content using the selected provider.
    
    Args:
        provider: 'gemini', 'groq', or 'cohere'
        api_key: API key for the provider
        questions: parsed questions list
        overview: deal overview text
        risks: parsed risks list
        rags: calculated RAG scores dict
        rfp_text: extracted RFP text
        
    Returns:
        dict: Parsed JSON with all PPT content fields
    """
    user_prompt = _build_user_prompt(questions, overview, risks, rags, rfp_text)
    
    callers = {
        'gemini': _call_gemini,
        'groq': _call_groq,
        'cohere': _call_cohere
    }
    
    if provider not in callers:
        raise ValueError(f"Unknown provider: {provider}")
    
    raw_text = callers[provider](api_key, SYSTEM_PROMPT, user_prompt)
    result = _parse_json_response(raw_text)
    
    # Validate required fields
    required = ['key_justification', 'deal_overview', 'positive_notes', 'dimensions',
                'assumptions', 'next_steps']
    for field in required:
        if field not in result:
            raise ValueError(f"AI response missing required field: {field}")
    
    return result
