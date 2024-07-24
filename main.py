import openai
import os
import docx
import pandas as pd
import time

# Set your OpenAI API key from environment variable
openai.api_key = os.getenv("OPENAI_API_KEY")

file_path = '/Users/panzhiheng1/PycharmProjects/pythonProject/seeU_project/0kpzmd6vk656m9sr8ro8dr04vf038afj.docx'
output_dir = '/Users/panzhiheng1/PycharmProjects/pythonProject/seeU_project/pic'
csv_path = '/Users/panzhiheng1/PycharmProjects/pythonProject/seeU_project/math_knowledge.csv'



# Function to read the .docx file and extract text and images
def read_docx(file_path, output_dir):
    doc = docx.Document(file_path)
    full_text = []
    images = []

    # Loop through each paragraph and check for inline shapes (images)
    for para in doc.paragraphs:
        full_text.append(para.text)

    # Loop through each part of the document to extract images
    image_counter = 0
    for rel in doc.part.rels:
        if "image" in doc.part.rels[rel].target_ref:
            img = doc.part.rels[rel].target_part.blob
            img_filename = os.path.join(output_dir, f'image_{image_counter}.png')
            with open(img_filename, 'wb') as img_file:
                img_file.write(img)
            images.append(img_filename)
            full_text.append(f'[Image: {img_filename}]')
            image_counter += 1

    return '\n'.join(full_text), images

# Read the content of the .docx file and extract images
full_text, images = read_docx(file_path, output_dir)

# Read the CSV file
csv_data = pd.read_csv(csv_path)

# Prepare the data for the API request
csv_data_str = csv_data.to_string(index=False)

# Construct the prompt with references to the images and text
prompt = f"Here is a list of math problems: {full_text} with images. Solve these problems and list out the sequance number of the needed knowledge points from this file: {csv_data_str} for each question. Only return the question and the sequance numbers."

# Split the prompt into chunks if necessary to handle large texts
def split_text_into_chunks(text, chunk_size=2000):
    return [text[i:i + chunk_size] for i in range(0, len(text), chunk_size)]

chunks = split_text_into_chunks(prompt)

# Function to send request to OpenAI API with retry mechanism
def send_request_with_retry(messages, retries=5, wait_time=5):
    for i in range(retries):
        try:
            response = openai.ChatCompletion.create(
                model="gpt-4o",
                messages=messages,
                max_tokens=1500  # Adjust the max tokens as needed
            )
            return response.choices[0].message['content'].strip()
        except openai.error.RateLimitError as e:
            print(f"Rate limit error: {e}. Retrying in {wait_time} seconds...")
            time.sleep(wait_time)
            wait_time *= 2  # Exponential backoff
        except Exception as e:
            print(f"An error occurred: {e}")
            break
    return None

# Process each chunk separately and accumulate responses
responses = []
accumulated_context = ""

for chunk in chunks:
    messages = [
        {"role": "system", "content": "You are a helpful assistant."},
        {"role": "user", "content": accumulated_context + chunk}
    ]

    # Send the request to the OpenAI API with retry mechanism
    response_text = send_request_with_retry(messages)
    if response_text:
        responses.append(response_text)
        accumulated_context += response_text + " "
    else:
        print("Failed to get a response after multiple retries.")
        break

# Print all responses
for response in responses:
    print(response)


print(response.choices[0].message['content'].strip())
