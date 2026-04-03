"""Quick test script for the Ingestion step using VN8910 data."""

import json
import sys
import io
import logging

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
logging.basicConfig(level=logging.INFO, format="%(levelname)s  %(message)s")

from surveyflow.steps.ingestion.step import IngestionStep

# ── Definition (VN8910 – Survey ID 722987) ─────────────────────────────────
DEFINITION = {
    "survey": {
        "survey_id": 722987,
        "title": "VN8910 - XX thanh trung",
        "english_title": "VN8910 - Sausage - pasteurized",
        "status": "2",
        "start_date": "2026-02-11 16:02:32",
        "end_date": "",
    },
    "questions": [
        {"question_id": 793092, "position": 1,  "question": "audio record",    "english_question": "audio record",    "type": 40,   "input_type": 40, "mandatory": True, "status": 1},
        {"question_id": 792916, "position": 2,  "question": "Ten shoppers",    "english_question": "What is your name?",                        "type": 1106, "input_type": 51, "mandatory": True, "status": 1},
        {"question_id": 792917, "position": 3,  "question": "SĐT",             "english_question": "What is your phone number?",                 "type": 1107, "input_type": 52, "mandatory": True, "status": 1},
        {"question_id": 792192, "position": 4,  "question": "Quota",           "english_question": "PVV only: Please select the right quota",   "type": 2,    "input_type": 0,  "mandatory": True, "status": 1},
        {"question_id": 792193, "position": 5,  "question": "Nganh nghe",      "english_question": "Do you or any of your family member living with you works in any of the following?", "type": 2, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 793000, "position": 6,  "question": "Dia diem",        "english_question": "Where do you live?",                        "type": 1109, "input_type": 1,  "mandatory": True, "status": 1},
        {"question_id": 792196, "position": 7,  "question": "Gioi tinh",       "english_question": "Please select your gender",                 "type": 2,    "input_type": 0,  "mandatory": True, "status": 1},
        {"question_id": 792197, "position": 8,  "question": "Thu nhap",        "english_question": "What is your monthly household income ?",   "type": 2,    "input_type": 0,  "mandatory": True, "status": 1},
        {"question_id": 792198, "position": 9,  "question": "Tuoi",            "english_question": "What is your age ?",                        "type": 2,    "input_type": 0,  "mandatory": True, "status": 1},
        {"question_id": 792199, "position": 10, "question": "Hon nhan",        "english_question": "Please choose your marital status",         "type": 2,    "input_type": 0,  "mandatory": True, "status": 1},
        {"question_id": 792200, "position": 11, "question": "So con",          "english_question": "How many children do you have?",            "type": 2,    "input_type": 0,  "mandatory": True, "status": 1},
        {"question_id": 792201, "position": 12, "question": "Tuoi con",        "english_question": "What are the ages of your children? Interviewer note: Please select one age per child (e.g. if the respondent has 3 children, select 3 age options).", "type": 3, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792290, "position": 13, "question": "Instruction",     "english_question": "PVV read out loud: How to distinguish between Sterilized & Pasteurized Sausage:\n\n\n\nPasteurized and sterilized sausages mainly differ in their shelf life, and storage conditions.\n\n\n\n- Pasteurized sausages (usually kept refrigerated with a short shelf life of about 2\u20133 weeks) \n\n- Sterilized sausages (can be stored at room temperature with a shelf life of several months)", "type": 31, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792202, "position": 14, "question": "Tan suat",        "english_question": "How often do you & your child/children consume pasteurized sausages nowadays?", "type": 2, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792203, "position": 15, "question": "Cach chon - with kids", "english_question": "Please tell me which one of the following statements BEST describes how you choose the sausage brands you & your family use.", "type": 2, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792204, "position": 16, "question": "Cach chon - no kids",   "english_question": "Please tell me which one of the following statements BEST describes how you choose the sausage brands you & your family use.", "type": 2, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792913, "position": 17, "question": "Instruction 2",   "english_question": "Interviewer Instruction: Next questions are about SAUSAGE CONSUMPTION BEHAVIOURS\n\n\n\nIf the child chooses the brand, the child should answer all remaining questions directly, under parental supervision.\n\nIf the parents choose  the brands,  the parent answers these question.\n\n ", "type": 31, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792205, "position": 18, "question": "TOM Awareness",  "english_question": "Which Pasterurized Sausage brand/variant are you aware of ", "type": 9, "input_type": 101, "mandatory": True, "status": 1},
        {"question_id": 792286, "position": 19, "question": "NET brands",     "english_question": "INTERVIEWERS ANSWER\n\nPlease select the NET brands in (Q14) in above Awareness question.\n\n\n\nShow brands selected in Q14\n\n\n\n{QQ14/792205/Selected}", "type": 3, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792206, "position": 20, "question": "P6M usage",      "english_question": "Which brand/variants have you used P6M (pipe from Q14)", "type": 9, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792207, "position": 21, "question": "MFU",            "english_question": "Which variant have you consumed most often in P6M (Pipe from Q 15)", "type": 8, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792208, "position": 22, "question": "Reasons consume","english_question": "Which of the following factors are your reasons for consuming pasteurised sausage? Rank minimum top 3", "type": 3, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792525, "position": 23, "question": "Ranking reasons","english_question": "Which of the following factors are your reasons for consuming pasteurised sausage?  Rank minimum top 3", "type": 6, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792210, "position": 24, "question": "Occasions",      "english_question": "During which occasions or activities do you typically consume pasteurised sausages?", "type": 3, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792211, "position": 25, "question": "Who eats",       "english_question": "Who in your household consumes pasteurised sausages ?", "type": 3, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792213, "position": 26, "question": "Where eat",      "english_question": "Where do you usually consume pasteurised sausages ? Select all that apply.", "type": 3, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792214, "position": 27, "question": "How eat",        "english_question": "How do you usually consume pasteurised sausage ?", "type": 3, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792215, "position": 28, "question": "Meat type",      "english_question": "Imagine your ideal pasteurised sausage. What is the meat type you would prefer? Select minimum 3.", "type": 3, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792624, "position": 29, "question": "Ranking meat",   "english_question": "Rank top 3 meat type of your ideal pasteurised sausage.", "type": 6, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792217, "position": 30, "question": "Texture",        "english_question": "Regarding texture, how do you prefer your pasteurised sausage ?", "type": 2, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792218, "position": 31, "question": "Texture surface","english_question": "And how do you prefer your pasteurised sausage ?", "type": 2, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792219, "position": 32, "question": "Sausage type",   "english_question": "What type of Pasteurized sausage do you prefer?", "type": 3, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792625, "position": 33, "question": "Ranking sausage","english_question": "Rank top 3 Pasteurized sausage you prefer. ", "type": 6, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792221, "position": 34, "question": "Flavour",        "english_question": "Below is a list of flavours. Which would you find appealing in pasteurized sausages. Select all that apply.", "type": 3, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792222, "position": 35, "question": "Intl cuisine",   "english_question": "You mentioned you find INTERNATIONAL CUISINE appealing in pasteurised sausages. Which of the following dishes do you find most appealing? Select all that apply.  ", "type": 3, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792223, "position": 36, "question": "Herbs spices",   "english_question": "You mentioned you find HERBS & SPICES appealing in sterilised pasteurised. Which of the following do you find most appealing? Select all that apply.", "type": 3, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792224, "position": 37, "question": "Spicy chili",    "english_question": "You mentioned you find SPICY CHILI & ROOTS appealing. Which flavours do you find most appealing? Select all that apply", "type": 3, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792225, "position": 38, "question": "Plant-based",    "english_question": "What would you say about plant-based sausage ? ", "type": 2, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792287, "position": 39, "question": "Instruction purchase", "english_question": "Interviewer Instruction: Next questions are about SAUSAGE PURCHASE BEHAVIOURS\n\n\n\nSHOW ANSWER OF Q13A : {QQ13A/792203/Selected}", "type": 2, "input_type": 0, "mandatory": True, "status": 0},
        {"question_id": 792288, "position": 40, "question": "Purchase factors","english_question": "Please tell me all the factors that you consider when purchasing pasteurised sausages", "type": 3, "input_type": 101, "mandatory": True, "status": 1},
        {"question_id": 792978, "position": 41, "question": "Nutrition open",  "english_question": "You have just answered \"Has nutritional value\" is 1 of the top consideration factor when purchase pasteurised sausage. So what nutritional value do you expect from pasteurised sausage?", "type": 1, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792289, "position": 42, "question": "Point allocation","english_question": "From the factors you selected previously, please allocate a total of 100 points across the options below based on their importance.\n\nThe higher the score, the more important the selected factor is.\n\nThe total score must equal 100 points.", "type": 1, "input_type": 100, "mandatory": True, "status": 1},
        {"question_id": 792227, "position": 43, "question": "Hinder purchase", "english_question": "Which of these reasons would hinder you purchasing Pasteurized sausages?", "type": 3, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792228, "position": 44, "question": "Pack check",      "english_question": "When buying pasteurised sausages what do you usually check on the pack? Select all that apply.", "type": 3, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792229, "position": 45, "question": "Claims",          "english_question": "What product claims will make you buy more Pasteurized sausages? Please select all that apply", "type": 3, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792230, "position": 46, "question": "Where buy",       "english_question": "Where do you normally buy your pasteurised sausages ?", "type": 3, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792231, "position": 47, "question": "Pack size",       "english_question": "What size pack do you usually buy Pasteurized sausage ? ", "type": 2, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792232, "position": 48, "question": "Spending",        "english_question": "How much do you spend on Pasteurized sausages per purchase on average?\n\n\n\n  {QQ30/792231/Selected}", "type": 2, "input_type": 0, "mandatory": True, "status": 1},
        {"question_id": 792999, "position": 49, "question": "reward",          "english_question": "reward", "type": 1, "input_type": 68, "mandatory": True, "status": 1},
    ],
}

# ── Load rows pages ─────────────────────────────────────────────────────────
BASE = (
    "C:/Users/PC/.claude/projects/"
    "C--Users-PC-OneDrive-DevZone-PyPackages-surveyflow--claude-worktrees-clever-boyd/"
    "c1b76edf-5bd6-4519-a4cb-f85c372c459f/tool-results/"
)
with open(BASE + "mcp-92b3762a-2e81-47a2-9b3e-095c0b4486d7-get_survey_rows-1775188592657.txt", encoding="utf-8") as f:
    page1 = json.load(f)
with open(BASE + "mcp-92b3762a-2e81-47a2-9b3e-095c0b4486d7-get_survey_rows-1775188639757.txt", encoding="utf-8") as f:
    page2 = json.load(f)

# ── Run ────────────────────────────────────────────────────────────────────
step = IngestionStep()
context = step.run({
    "definition": DEFINITION,
    "rows_pages": [page1, page2],
    "output_dir": "C:/Users/PC/OneDrive/DevZone/PyPackages/surveyflow/output_test",
})

df = context["rawdata"]
meta = context["metadata"]

print(f"\n{'='*60}")
print(f"rawdata shape  : {df.shape}  ({df.shape[0]} respondents × {df.shape[1]} columns)")
print(f"Saved CSV      : {context['rawdata_path']}")
print(f"Saved JSON     : {context['metadata_path']}")

print(f"\n--- Columns ---")
print(list(df.columns))

print(f"\n--- Sample row (respondent 1) ---")
r = df.iloc[0]
for col in ["task_id", "date_time", "profile_status",
            "q4", "q6", "q7", "q8", "q9",
            "q18", "q20", "q21", "q22", "q42"]:
    print(f"  {col:20s}: {r.get(col, 'N/A')[:80] if r.get(col) else '(empty)'}")

print(f"\n--- metadata.questions sample (q7, q18, q22) ---")
for q in meta["questions"]:
    if q["label"] in ("q7", "q18", "q22"):
        print(f"  {q['label']:8s} | {q['answer_type']:15s} | values({len(q['values'])}): {dict(list(q['values'].items())[:5])}")
