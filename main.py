import docx
import pandas as pd
import openai
import os

file_path = '/Users/panzhiheng1/PycharmProjects/pythonProject/seeU_project/0kpzmd6vk656m9sr8ro8dr04vf038afj.docx'
file_path1 = '/Users/panzhiheng1/PycharmProjects/pythonProject/seeU_project/math_knowledge.csv'
openai.api_key = os.getenv("OPENAI_API_KEY")
output_directory_path = '/Users/panzhiheng1/PycharmProjects/pythonProject/seeU_project/pic'

def read_docx(file_path):
    doc = docx.Document(file_path)
    text = []
    for para in doc.paragraphs:
        text.append(para.text)
    for rel in doc.part.rels:
        if "image" in doc.part.rels[rel].target_ref:
            img = doc.part.rels[rel].target_part.blob
            img_filename = os.path.join(output_directory_path, os.path.basename(doc.part.rels[rel].target_ref))
            with open(img_filename, 'wb') as img_file:
                img_file.write(img)
            text.append(f'[Image: {img_filename}]')

    return '\n'.join(text)


full_text = read_docx(file_path)
knowledge_points = pd.read_csv(file_path1)
messages = [
    {"role": "system", "content": "You are a helpful assistant."},
    {"role": "user",
     "content": f"For each question in the question list, provid the sequance number of needed knowledge point to solve the question. Here is a list of Chinese Math questions: {full_text} Here is a list of knowledge points: {knowledge_points}  "}
]

response = openai.ChatCompletion.create(
    model = "gpt-4",
    messages=messages,
    max_tokens=3000  # Adjust the max tokens as needed
)


print(response.choices[0].message['content'].strip())
