# from docx import Document

# data = {
#     '[Name]': 'John Doe',
#     '[Title]': 'Software Developer',
#     '[Education]': 'Bachelor of Science in Computer Science',
#     # Add more data as needed
# }


# doc = Document("template.docx")

# for paragraph in doc.paragraphs:
#     for run in paragraph.runs:
#         print(paragraph)

# for key, value in data.items():
#     for paragraph in doc.paragraphs:
#         if key in paragraph.text:
#             paragraph.text = paragraph.text.replace(key, str(value))

# # Save the filled-out document
# doc.save("test1.docx")

import sys
import openai
import json
import os

def cv_reader(cv):
    # Set OpenAI API key & Initialize GPT Client
    openai.api_key = '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'


    # Initialize Chatbot Personality
    conversation = [
    {"role": "system", "content": """I will pass you a CV template divided into sections with [SECTION xyz] subheaders. I want you to read the Instructions mentioned at the bottom of the template such that I can pass you another actual CV and you can format it according to the template and give me back a list as specified in the instructions. Do NOT invent your own sections, regardless of whether the CV i give you has extra content, stick to the 5 sections mentioned in the template. I want you to pass me a LIST object suitable for interpretation in Python:

[SECTION 1A] NAME
[SECTION 1B] Location | LinkedIn| Phone | Email
[SECTION 2] EDUCATION
University                                                                                                                                         Location
Major/Course                                                                                                                                          Dates                                   
•	GPA[if 3.2 or above], Organizations, Coursework, etc.                                                                                                                                                   

[SECTION 3] WORK EXPERIENCE 
Company                                                                                                                                          Location
Position                                                                                                                                                   Dates
•	This section regarding experiences has bulleted accomplishments, which provide examples of when you successfully used the skills employers are seeking. Make sure you have between 2 and 6 bullet points in each section. 
•	Your bullet points should start with a strong action verb, which then follows with an explanation of what you were doing, describing how you did it, and most importantly if applicable, any achievements. Statements should convey your strengths/proficiencies in one or more skills that intrigue the employer by showing examples of when you have used them. 
•	When writing about your experience, consider these questions: What was the result/outcome of your work? What were your accomplishments? How did you impact the organization? What skills/knowledge did you grow? How does this experience relate to your internship/employment goal?
Company                                                                                                                                                                        Location
Position                                                                                                                                                                                    Dates
•	Your bullet statements should be in proper tense, using-ed for past experiences and present tenses for current positions. Make sure that your writing is free of grammatical errors and punctuation. · 
•	When including numerical achievements during your experiences, make sure to include (if applicable) the quantity, population, frequency, and impact of your work whenever possible. 
•	To make your resume flow, read it over. Check and see if it is easy to read with no overflowing of text. You should avoid the usage of different colors, multiple fonts, pictures, and brief/too dense information. Your resume should show who you are while being professional.

[SECTION 4] LEADERSHIP EXPERIENCE (OPTIONAL)
Company                                                                                                                                          Location
Position                                                                                                                                                   Dates
•	This section is optional if you have various leadership experiences and other activities you want employers to know. By having multiple sections, it allows you to emphasize your most relevant experience. 
•	Positions within this section should be formatted similarly to previous experience sections, including bullet points if necessary. 
•	You may also include work experiences that may not be directly related to the job/internship you are applying to but add to your credibility by exemplifying your past work experiences.

[SECTION 5] SKILLS & ADDITIONAL TRAINING
Skills: These skills should be concrete and testable. These should not be soft skills like communication, organizational, and interpersonal skills, but instead incorporated into your bulleted accomplishment statements above. Interests: What type of additional training did you perform, or certificates did achieve? 

Instructions:

Return a list [SECTION 1A, SECTION 1B, SECTION 2, SECTION 3, SECTION 4, SECTION 5]

Where:
SECTION 1A is Name

SECTION 1B should be a list [Location, LinkedIn, Phone, Email]. If one of these is missing, just input a placeholder for example LOCATION.

SECTION 2 should be a list of education. The first index represents how many different education experiences there are and the following indices contain those work experiences in sublists. Each Education sublis should be of the form [University/School Name, Location, Major/Course, Dates, [a list containing GPA, organizations, coursework, etc]]. If there are none, make the first index 0.

SECTION 3 should be a list of work experiences. The first index represents how many different work experiences there are and the following indices contain those work experiences in sublists. Each Leadership experience sublist should be of the form [Company, Location, Position, Dates, Description of experience]

SECTION 4 should be a list of leadership experiences. The first index represents how many different leadership experiences there are and the following indices contain those leadership experiences in sublists. Each Leadership experience sublist should be of the form [Company, Location, Position, Dates, Description of experience]

SECTION 5 should be a list of skills and additional training where the first index represents how many different skills and training there are and the following indices contain those additional skill and additional training bullet points. If there are none, make the first index 0.

Remember, when I pass you the actual CV, just give me the list and nothing else. Just  the list object without a name and without any code, I will manage the code."""},
    ]


    # Send in New CV
    conversation.append({"role": "user", "content": user_message})

    # Get model-generated response
    response = chat_with_model(conversation)
    conversation.append({"role": "assistant", "content": response})

    
    # Update Conversation History
    with open(file_path, "w") as json_file:
        json.dump(conversation, json_file)

    # Write the output to standard output
    sys.stdout.write(response)
    # Ensure the output is flushed
    sys.stdout.flush()

    with open(file_path, "w") as json_file:
        json.dump(conversation, json_file)


# Function to create a conversation and get model-generated responses
def chat_with_model(messages):
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=messages
    )
    return response.choices[0].message["content"]

def get_visa_info(messages):
    messages.append({"role": "user", "content": "What are the visa requirements for me to travel from my city of origin to my city of destination with my given passport?"})
    visa_reqs = chat_with_model(messages)
    return visa_reqs

list = [['Khamis Mohammad Hamad Almazrouei'],
        ['Liwa, Al Dhafra Region','LinkedIn','+971501551120','khamisliwa9@gmail.com'],
 ['2', ['Institute of Management Technology, Dubai', 'Dubai', 'Applied Bachelor Science in Business Administration', '2017 – 2022', ['GPA 2.99', 'Major Courses:', 'Global Logistics Management', 'Advance Managerial Accounting']]], 
 '0', [], 
 '0', [], 
 '0', [], 
 '1', ['Redcresent, Madinat Zayed', '1 day', 'March 2017', 'Did filing, data feeding'], 
 '1', ['Ministry of culture and Knowledge development- Madinat Zayed', '2 weeks, July 2017', 'Well-handled emails and phone calls for the summer cultural event'], 
 '1', ['Ministry of culture and Knowledge development- Madinat Zayed', '1 Week', 'August 2018', 'Printing and files arrangements'], 
 '1', ['Emirates Foundation Forum', 'One day', 'Sep 2018', 'Printing and files arrangements'], 
 '1', ['ADSSC-Abu Dhabi', '1 year 2014-2015', 'Customer survives'], 
 ['2', ['Good in the MS-Office tools which includes MS-Word, MS-Excel and MS-Power point', 'Good communication skills that helps me to make friends with the unknown people and gain new relationship', 'I have developed good skills in facing the challenges of the real-world and finds the solutions to solve them', 'Confident, articulate in dealing with tasks', 'I demonstrate good teamwork', 'I am capable of looking for innovative solutions to the problems and I am highly self-motivated', 'Good in management of time', 'Arabic as native, good in English( Speak, Write, understand) and can speak Hindi language']]]