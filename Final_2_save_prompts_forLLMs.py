import re
import openai
import pandas as pd
import os

# Load students' answers from the file
file_path = 'FINAL TXT (1-73).xlsx'
df = pd.read_excel(file_path, sheet_name = 'Sheet1')

# Set your OpenAI API key
openai.api_key = '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'

# Define the path to the prompts folder
script_dir = os.path.dirname(os.path.abspath(__file__))
prompts_dir = os.path.join(script_dir, 'prompts\\Q2\\')

# Load the question and grading details from the prompts folder
with open(os.path.join(prompts_dir, 'the_question.txt'), 'r') as f:
    the_question = f.read()

with open(os.path.join(prompts_dir, 'grading_prompt.txt'), 'r') as f:
    grading_prompt = f.read()

with open(os.path.join(prompts_dir, 'possible_deductions.txt'), 'r') as f:
    possible_deductions = f.read()

with open(os.path.join(prompts_dir, 'after_evaluation.txt'), 'r') as f:
    after_evaluation = f.read()

# Function to evaluate all parts of a student's answers in a single API call
def evaluate_answers_combined(answers):
    # Combine answers for all parts into a single string
    answers_combined = f"""
    Question 2a:
    {answers['2a']}
    
    Question 2b:
    {answers['2b']}
    
    """
    
    # Create the API prompt
    prompt = grading_prompt.format(the_question, possible_deductions, after_evaluation) + "\n\nStudent Answers:\n" + answers_combined
    print(prompt)

    
    # Send the request to OpenAI API
    response = openai.ChatCompletion.create(
        model="gpt-4-turbo",
        messages=[
            {"role": "system", "content": "You are a teaching assistant evaluating a student's answers about Object-Oriented Programming concepts in C++ or Java for correctness and completeness."},
            {"role": "user", "content": prompt}
        ]
    )
    
    # Extract the response text
    evaluation_result = response['choices'][0]['message']['content']
    
    # Define regex patterns to extract grades and feedback for each part
    grades = {}
    feedback = {}
    for part in ['2a', '2b']:
        # Attempt to extract grades
        grade_match = re.search(rf"Grade_{part}:\s*(\d+)", evaluation_result)
        grades[part] = int(grade_match.group(1)) if grade_match else None
        
        # Attempt to extract feedback
        feedback_match = re.search(rf"{part}:\s*(.*?)(?:Grade|Final Comments)", evaluation_result, re.DOTALL)
        feedback[part] = feedback_match.group(1).strip() if feedback_match else "Feedback missing"
    
    return evaluation_result, grades, feedback

# Function to save the generated prompt instead of sending it to the ChatGPT API
# Function to save the generated prompt to a file with STD_CODE in the filename
def save_prompt_to_file(prompt, std_code):
    output_file = f'savedprompts2\generated_prompt_{std_code}.txt'
    with open(output_file, 'w', encoding='utf-8') as file:
        file.write(prompt)
    print(f"Prompt saved to {output_file}.")

# Function to evaluate all parts of a student's answers in a single API call (modified to save prompt)
def evaluate_answers_combined_saveprompt(answers, std_code):
    # Combine answers for all parts into a single string
    answers_combined = f"""
    Question 2a:
    {answers['2a']}
    
    Question 2b:
    {answers['2b']}
    
    """
    
    # Create the API prompt
    prompt = grading_prompt.format(the_question, possible_deductions, after_evaluation) + "\n\nStudent Answers:\n" + answers_combined
    
    # Save the prompt to a file instead of sending it to the ChatGPT API
    save_prompt_to_file(prompt, std_code)

    return None, {}, {}  # Return placeholders for evaluation, grades, and feedback




# Process each student's answers
results = []

for index, row in df.iterrows():
    student = {
        'Name': row['Name'],
        'Email': row['Email'],
        'STD_CODE': row['STD_CODE']
    }
    
    # Extract answers for each part
    answers = {
        '2a': row['Write the answer to question 2.a directly here.'],
        '2b': row['Write the answer to question 2.b directly here.']
    }
    
    # Evaluate all parts in a single API call
    evaluation, grades, feedback = evaluate_answers_combined_saveprompt(answers, student['STD_CODE'])

 
    
    
