# MCQ Master — Streamlit MCQ Quiz App

A simple Streamlit application for conducting MCQ quizzes from Word (`.docx`) files.

## Features

- Upload a `.docx` MCQ file.
- Reads MCQs from Word paragraphs and tables.
- Supports questions with A, B, C and D options.
- Supports answer formats such as `Answer: B`, `Ans: B`, and `Correct Answer: B`.
- Students select one answer for each question.
- Calculates:
  - Correct answers
  - Wrong answers
  - Unanswered questions
  - Percentage score
- Optional question shuffling.
- Optional answer review.
- Retake quiz button.
- Responsive Streamlit UI.

## Word document format

Use this format:

```text
1. What does MATLAB stand for?
A. Mathematical Laboratory
B. Matrix Laboratory
C. Machine Laboratory
D. Matrix Language
Answer: B

2. Which command is used to plot a 2D graph?
A. stem
B. subplot
C. plot
D. size
Answer: C
```

The supplied Signals and Systems and DSP MCQ documents follow this general structure, so they can be used as test files.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Create a GitHub repository.
2. Upload:
   - `app.py`
   - `requirements.txt`
3. Open Streamlit Community Cloud.
4. Select your GitHub repository.
5. Select `app.py` as the main file.
6. Deploy.

No API key is required because this version parses existing MCQs from the uploaded Word document.

## Important clarification

This version is a **Word-to-Quiz parser**: the Word document must already contain MCQs and their correct answers.

If you want the next version to take a **lecture/manual without MCQs and automatically generate new MCQs using AI**, an AI API (for example, Groq/OpenAI) can be added as a separate feature.
