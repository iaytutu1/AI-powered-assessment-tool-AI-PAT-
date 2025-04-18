import re
import openai
import pandas as pd
import os

# Load students' data from the new Excel file
file_path = 'Final_Exam_Merged_Appeals_Last5std.xlsx'
df = pd.read_excel(file_path)

# Clean column headers
df.columns = [col.strip() for col in df.columns]

df.rename(columns={
    "Write your objections in detailed (in English).": "Appeal",
}, inplace=True)


openai.api_key = '' 

# Define base directory containing question and rubric files
base_dir = 'prompts\\'  # Directory shown in the image
question_folders = ['Q1', 'Q2', 'Q3', 'Q4']

# Function to load question and grading prompts from folders
def load_questions_and_rubrics(base_dir, question_folders):
    combined_question = []
    combined_rubric = []
    
    for folder in question_folders:
        folder_path = os.path.join(base_dir, folder)
        question_file = os.path.join(folder_path, 'the_question.txt')
        rubric_file = os.path.join(folder_path, 'grading_prompt.txt')
        
        if os.path.exists(question_file) and os.path.exists(rubric_file):
            with open(question_file, 'r') as f:
                question = f.read()
            with open(rubric_file, 'r') as f:
                rubric = f.read()
            
            combined_question.append(f"{folder} Question:\n{question}")
            combined_rubric.append(f"{folder} Rubric:\n{rubric}")
        else:
            print(f"Missing question or rubric file in {folder_path}")
    
    return "\n\n".join(combined_question), "\n\n".join(combined_rubric)

# Load all questions and rubrics
combined_question, combined_rubric = load_questions_and_rubrics(base_dir, question_folders)

# Function to concatenate columns dynamically
# Function to concatenate columns dynamically, converting non-string values
def concatenate_columns_dynamic(df, keywords, new_column_name):
    matching_cols = [col for col in df.columns if any(keyword in col for keyword in keywords)]
    if matching_cols:
        df[new_column_name] = df[matching_cols].astype(str).fillna('').apply(lambda x: ' '.join(x).strip(), axis=1)

# Concatenate Evaluation, Answer, Appeal, and Feedback columns
concatenate_columns_dynamic(df, ['Evaluation', 'Feedback'], 'Evaluation Combined')
concatenate_columns_dynamic(df, ['Write the answer to'], 'Answer Combined')
concatenate_columns_dynamic(df, ['Appeal'], 'Appeal Combined')

# Function to review appeals
def review_appeal(student_row):
    std_code = student_row['STD_CODE']
    evaluation = student_row.get('Evaluation Combined', '')
    appeal = student_row.get('Appeal Combined', '')
    answer = student_row.get('Answer Combined', '')

    # Prepare the GPT prompt
    prompt = (
        f"Here are the exam questions:\n{combined_question}\n\n"
        f"Here is the grading rubric:\n{combined_rubric}\n\n"
        f"Student's Combined Answer:\n{answer}\n\n"
        f"Combined Evaluation and Feedback:\n{evaluation}\n\n"
        f"Appeal from the student:\n{appeal}\n\n"
        f"Instructions: Based on the combined answer, appeal, rubric, and provided feedback, determine if the grade should increase."
        f" Provide reasons for any increase, but do not decrease the grade."
        f" Give the exact points for each reason and the total adjustment in your response."
    )
    print(prompt)
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4-turbo",
            messages=[
                {"role": "system", "content": "You are a teaching assistant reviewing student appeals."},
                {"role": "user", "content": prompt}
            ]
        )

        appeal_result = response['choices'][0]['message']['content']
        grade_match = re.search(r"NewGrade:\s*([\d\.]+)", appeal_result)
        new_grade = float(grade_match.group(1)) if grade_match else student_row.get('TOTAL', 0)

        return appeal_result, new_grade

    except Exception as e:
        print(f"Error during appeal review: {e}")
        return None, student_row.get('TOTAL', 0)

# Process appeals for all students
results = []

for index, row in df.iterrows():
    appeal_result, new_grade = review_appeal(row)
    results.append({
        'STD_CODE': row['STD_CODE'],
        'Original Grade': row.get('TOTAL', 0),
        'New Grade': new_grade,
        'Appeal Result': appeal_result
    })


# Save the results to an Excel file
results_df = pd.DataFrame(results)
output_file = 'appeal_review_results_Last5std.xlsx'
results_df.to_excel(output_file, index=False)
print(f"Appeal review completed and saved to {output_file}.")
