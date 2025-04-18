import re
import openai
import pandas as pd
import os

# Load students' answers from the file
file_path = 'FINAL TXT (1-73).xlsx'
df = pd.read_excel(file_path, sheet_name = 'Sheet1')

openai.api_key =""

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

# Retry mechanism for students with missing grades
def retry_missing_grades(student, answers, evaluation, grades, feedback, max_retries=1):
    retries = 0
    while retries < max_retries:
        missing_parts = [part for part, grade in grades.items() if grade is None]
        if not missing_parts:
            break  # All grades are present, no need to retry
        
        print(f"Retrying for {student['Name']}... Missing grades: {missing_parts}")
        # Re-evaluate answers for missing parts
        retry_evaluation, retry_grades, retry_feedback = evaluate_answers_combined(answers)
        
        # Update missing parts with re-evaluated grades and feedback
        for part in missing_parts:
            grades[part] = retry_grades.get(part, grades[part])  # Update only if found
            feedback[part] = retry_feedback.get(part, feedback[part])
        
        retries += 1
    
    if any(grade is None for grade in grades.values()):
        print(f"Unresolved missing grades for {student['Name']} after {max_retries} retries.")
    
    return grades, feedback

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
    evaluation, grades, feedback = evaluate_answers_combined(answers)
    
    # Retry for missing grades
    grades, feedback = retry_missing_grades(student, answers, evaluation, grades, feedback)
    
    # Append the results to the list
    results.append({
        'Name': student['Name'],
        'Email': student['Email'],
        'STD_CODE': student['STD_CODE'],
        'Evaluation': evaluation,
        'Feedback_2a': feedback.get('2a'),
        'Grade_2a': grades.get('2a'),
        'Feedback_2b': feedback.get('2b'),
        'Grade_2b': grades.get('2b')
    })

# Convert results to a DataFrame for easy saving and export
results_df = pd.DataFrame(results)
output_file = 'graded_students_answers_FinalExam_with_retries_Q2_secondtime.xlsx'
results_df.to_excel(output_file, index=False)
print(f"Grading completed and saved to {output_file}.")