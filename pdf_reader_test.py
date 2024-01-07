from PyPDF2 import PdfReader

import openai

pdf_reader = PdfReader('Cv.M7md.pdf')

num_pages = pdf_reader.getNumPages()

text = ""

for x in range(num_pages):
  page = pdf_reader.pages[0]

  text = text + " " + pages.extract_text()

api_key = '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'

openai.api_key = api_key

file = client.files.create(
   file = open("template.docx")
   purpose = "assistants"
)

# Function to create a conversation and get model-generated responses
def chat_with_model(unformatted_cv_text):
    
  completion = client.chat.completions.create(
    model="gpt-4-1106-preview",
    tools = [{"type": "retrieval"}]
    messages=[
      {"role": "system", "content": """You are assisting CV conversion to ATS friendly format. You are in the middle stage in a pipeline where a CV is retrieved from a database and its text contents are passed on to you; you then fill in the CV template found in template.docx.

      Template.docx contains 4 main sections, 
      Section 1: The name, contact info, and location header

      Section 2: Education

      Section 3: Work Experience (If found on original CV)

      Section 4: Leadership Experience (if any, if not found leave this blank, else if you find anything relating to leadership format it in a formal manner and input it here)

      Section 5: Skills and Additional Learning

      When you format a CV you must ensure the date and name spacings are kept from template.docx"""
        },
      {"role": "user", "content": unformatted_cv_text}
    ],
      
    )

    # print(completion.choices[0].message.content)
    return response.choices[0].message["content"]

# Create a conversation with user and assistant messages
conversation = [
    {"role": "system", "content": """You are assisting CV conversion to ATS friendly format. You are in the middle stage in a pipeline where a CV is retrieved from a database and its text contents are passed on to you; you then fill in the following CV template:
    (Section 1 start) NAME
Location | LinkedIn| Phone | Email
EDUCATION
University                                                                                                                                         Location
Major/Course                                                                                                                                          Dates                                   
•	GPA[if 3.2 or above], Organizations, Coursework, etc.                                                                                                                                                   
(Section 1 end)
     
(Section 2 start)
WORK EXPERIENCE 
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
(Section 2 end)
     
(Section 3 start)
LEADERSHIP EXPERIENCE (OPTIONAL)
Company                                                                                                                                          Location
Position                                                                                                                                                   Dates
•	This section is optional if you have various leadership experiences and other activities you want employers to know. By having multiple sections, it allows you to emphasize your most relevant experience. 
•	Positions within this section should be formatted similarly to previous experience sections, including bullet points if necessary. 
•	You may also include work experiences that may not be directly related to the job/internship you are applying to but add to your credibility by exemplifying your past work experiences.
(Section 3 end)
     
(Section 4 start)
SKILLS & ADDITIONAL TRAINING
Skills: These skills should be concrete and testable. These should not be soft skills like communication, organizational, and interpersonal skills, but instead incorporated into your bulleted accomplishment statements above. Interests: What type of additional training did you perform, or certificates did achieve? 
(Section 4 end)

Note: In the next message you will be passed an unformatted CV who's contents you have to translate into this provided template.
"""},
]


# Extend the conversation with a new user message
user_message = text
conversation.append({"role": "user", "content": user_message})

# Get model-generated response
response = chat_with_model(conversation)
print("Model's Response:", response)

conversation.append({"role": "assistant", "content": response})

# Test query for each section                  
user_message = "Provide me with section 1 of the CV"
conversation.append({"role": "user", "content": user_message})
response = chat_with_model(conversation)
print("Model's Response:", response)      
conversation.append({"role": "assistant", "content": response})

user_message = "Provide me with section 2 of the CV"
conversation.append({"role": "user", "content": user_message})
response = chat_with_model(conversation)
print("Model's Response:", response)
conversation.append({"role": "assistant", "content": response})

user_message = "Provide me with section 3 of the CV"
conversation.append({"role": "user", "content": user_message})
response = chat_with_model(conversation)
print("Model's Response:", response)
conversation.append({"role": "assistant", "content": response})

user_message = "Provide me with section 4 of the CV"
conversation.append({"role": "user", "content": user_message})
response = chat_with_model(conversation)
print("Model's Response:", response)
conversation.append({"role": "assistant", "content": response})

[
    "Khamis Mohammad Hamad Almazrouei",
    
    ["Liwa, Al Dhafra Region", "LinkedIn", "+971501551120", "khamisliwa9@gmail.com"],
    
    [
        2, 
        ["Institute of Management Technology, Dubai", "", "Applied Bachelor Science in Business Administration", "2017 - 2022", ["GPA 2.99", []]],
        ["Farouq School, Liwa", "", "Higher School Certificate", "2013", ["Marks: 79%", []]]
    ],
    
    [
        1, 
        ["Tadbeer", "", "Customer service", "2020-2022", ""]]
    ],
    
    [
        2, 
        ["Redcresent, Madinat Zayed", "", "Volunteering activities", "1 day, March 2017", "Did filing, data feeding"], 
        ["Ministry of culture and Knowledge development- Madinat Zayed", "", "Volunteering activities", "2 weeks, July 2017", "Well-handled emails and phone calls for the summer cultural event"],
        ["Ministry of culture and Knowledge development- Madinat Zayed", "", "Volunteering activities", "1 Week, August 2018", "Printing and files arrangements"],
        ["Emirates Foundation Forum", "", "Volunteering activities", "One day, Sep 2018", "Printing and files arrangements"],
        ["ADSSC-Abu Dhabi", "", "Volunteering activities", "1 year, 2014-2015", "Customer services"]
    ],

    [
        2,
        "MS-Office tools proficiency",
        "Good communication skills",
        "Problem-solving skills",
        "Confident and articulate",
        "Teamwork",
        "Innovative solutions",
        "Time management",
        "Arabic (native), English (Speak, Write, understand), Hindi (spoken)"
    ],

    # Comments:
    # - In the education section, the university and school names are given, but the locations are missing. Hence, the second entry in the sublist is an empty string.
    # - The GPA and coursework information is not provided, so the sublist for that is empty.
    # - Similarly, in the work experience section, the positions are missing, so those entries are empty strings.
    # - The leadership experience section is not provided, so the first index is 0.
    # - The section for skills and additional training is not provided, so the first index is 0.
    # - Ensure the consistency of sublist structure across all sections for accurate indexing.
    # - Follow the specified format strictly to facilitate proper interpretation in Python.
]


[
    "Khamis Mohammad Hamad Almazrouei",
    ["Liwa, Al Dhafra Region", "LinkedIn", "+971501551120", "khamisliwa9@gmail.com"],
    [
        2,
        ["Institute of Management Technology, Dubai", "", "Applied Bachelor Science in Business Administration", "2017 – 2022", ["GPA 2.99", "Major Courses: Global Logistics Management, Advance Managerial Accounting"]],
        ["Farouq School, Liwa", "", "Higher School Certificate", "2013", ["Marks: 79%"]]
    ],
    [
        0
    ],
    [
        1,
        ["Redcresent, Madinat Zayed", "Madinat Zayed", "1 day, March 2017", "Did filing, data feeding"],
        ["Ministry of culture and Knowledge development- Madinat Zayed", "Madinat Zayed", "2 weeks, July 2017", "Well-handled emails and phone calls for the summer cultural event"],
        ["Ministry of culture and Knowledge development- Madinat Zayed", "Madinat Zayed", "1 week, August 2018", "Printing and files arrangements"],
        ["Emirates Foundation Forum", "", "One day, Sep 2018", "Printing and files arrangements"],
        ["ADSSC-Abu Dhabi", "", "1 year, 2014-2015", "Customer services"]
    ],
    [
         8,
        "MS-Office tools: MS-Word, MS-Excel, and MS-PowerPoint",
        "Good communication skills and building relationships",
        "Problem-solving and innovative solutions",
        "Confident and articulate in dealing with tasks",
        "Good teamwork skills",
        "Self-motivated",
        "Time management",
        "Languages: Arabic (native), English (speaking, writing, understanding), Hindi (speaking)"
    ]
]

[
    "Khamis Mohammad Hamad Almazrouei",  
    
    ["Liwa, Al Dhafra Region", "LinkedIn", "+971501551120", "khamisliwa9@gmail.com"],  
    
    [
        1,  
        ["Institute of Management Technology, Dubai", "Dubai", "Applied Bachelor Science in Business Administration", "2017-2022", ["GPA 2.99", "Major Courses: Global Logistics Management, Advance Managerial Accounting"]]
    ],  
    
    [
        1,  
        ["Redcresent, Madinat Zayed", "Madinat Zayed", "Volunteering activities", "March 2017", "Did filing, data feeding"],
        ["Ministry of culture and Knowledge development- Madinat Zayed", "Madinat Zayed", "Volunteering activities", "July 2017", "Well-handled emails and phone calls for the summer cultural event"],
        ["Ministry of culture and Knowledge development- Madinat Zayed", "Madinat Zayed", "Volunteering activities", "August 2018", "Printing and files arrangements"],
        ["Emirates Foundation Forum", "", "Volunteering activities", "Sep 2018", "Printing and files arrangements"],
        ["ADSSC-Abu Dhabi", "Abu Dhabi", "Volunteering activities", "2014-2015", "Customer service"]
    ], 
    
    [
        1,  
        "Good skills in MS-Office tools (MS-Word, MS-Excel, MS-Powerpoint)",
        "Good communication skills",
        "Good problem-solving skills",
        "Confident and articulate in dealing with tasks",
        "Good teamwork skills",
        "Innovative and highly self-motivated",
        "Good time management skills",
        "Fluent in Arabic and English, and speaks Hindi"
    ],

    # Comments:
    # - The "Interests and Hobbies" section is not included in the resulting list as it is not one of the specified sections.
    # - The references section is also not included as it is not one of the specified sections.
    # - The work experience at Abu Dhabi Sewerage Service company is not included as it does not follow the specified format.
]