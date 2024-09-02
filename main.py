import os
import tkinter as tk
from tkinter import messagebox
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email import policy
from email.generator import BytesGenerator

def create_email(subject, body, to_email, from_email):
    """
    Create a MIMEText email object with the provided details.
    
    Parameters:
    subject (str): The subject of the email.
    body (str): The body content of the email.
    to_email (str): The recipient's email address.
    from_email (str): The sender's email address.
    
    Returns:
    MIMEMultipart: The email object ready to be sent or saved as a draft.
    """
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    return msg

def save_as_draft(email_obj, directory="Drafts", filename="draft.eml"):
    """
    Save the email as a draft in the specified directory.
    
    Parameters:
    email_obj (MIMEMultipart): The email object created by create_email().
    directory (str): The directory where the draft should be saved.
    filename (str): The filename to save the draft as (default: 'draft.eml').
    
    Returns:
    str: The full path to the saved draft file.
    """
    # Ensure the directory exists
    if not os.path.exists(directory):
        os.makedirs(directory)

    # Create the full file path
    full_path = os.path.join(directory, filename)

    # Save the email as a .eml file
    with open(full_path, 'wb') as f:
        BytesGenerator(f, policy=policy.default).flatten(email_obj)
    
    print(f"Draft saved as {full_path}")

    return full_path


from langchain_core.prompts import ChatPromptTemplate
from langchain_ollama.llms import OllamaLLM

template = """Question: {question}"""

prompt = ChatPromptTemplate.from_template(template)

model = OllamaLLM(model="llama3.1")

chain = prompt | model

content = "I have an event named Tech Appetizup 2024. It aims to Foster Youth's Dispositions in Innovation and Technology Development. There will be a Mobilization seminar on 12 July 2024, Hackathon on 13-14, 27-28 July 2024, Greater Bay Area Tours held from August 2024 to January 2025 and the Hong Kong Regional Competition on 25 March 2025. It is open to youths interested in Fintech, Education tech, Climate change tech. 3-5 people can apply as a team."


import pandas as pd


if __name__ == "__main__":
    # Specify the path to your Excel file
    file_path = 'temp_email_list.xlsx'

    # Read the Excel file
    df = pd.read_excel(file_path)

    # Loop through each row in the DataFrame
    for index, row in df.iterrows():
        temp = content

        # Combine the information into a single string
        email = row['Email']
        name = row['Name']
        school = row['College/High School']
        subject = row['Subject']
        industry = row['Industry']

        temp += f"my friend name is {name}, he/she is currently studying in {school}, studying {subject}, and is working in the {industry}" + "please write me a personalised email to him/her" + "Please sign the email with my name Nelson, Founder and Managing Director of Appetizup"

        email_content = chain.invoke({"question": temp})

        subject = email_content[email_content.find("Subject: ")+9:email_content.find('Dear')-2]

        email_content = email_content[email_content.find("Dear"):]

        to_email = email
        from_email = "nelson@outlook.com"

        # Create the email object
        email_obj = create_email(subject, email_content, to_email, from_email)

        # Save the draft to a dedicated location and show a popup
        dedicated_directory = "/Users/boscofungg/Desktop/AI project/ESG Nelson/Email Drafts"  # Customize this path as needed
        draft_path = save_as_draft(email_obj, directory=dedicated_directory, filename=f"{name}_draft.eml")
        print(f"Draft saved as {draft_path}")