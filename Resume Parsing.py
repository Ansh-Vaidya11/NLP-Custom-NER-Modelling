import docx2txt
import re
import string
import spacy
from spacy.matcher import Matcher
from spacy.tokens import Span

# Path to the downloaded language model folder
spacy_model_path = '/Users/Ansh/Downloads/en_core_web_sm-3.7.1/en_core_web_sm/en_core_web_sm-3.7.1'
nlp = spacy.load(spacy_model_path)

def text_preprocessing(text):
    # Removing extra spaces and empty lines
    cleaned_text = re.sub(r'\s+', ' ', text).strip()
    
    # Convert to lowercase
    cleaned_text = cleaned_text.lower()

    # Remove URLs
    cleaned_text = re.sub(r'https?://\S+|www\.\S+', '', cleaned_text)

    # Remove Punctuations and Special Characters (except for '.', '@', and '%')
    all_punctuation = string.punctuation
    exclude = ''.join(char for char in all_punctuation if char not in ['.', '@', '%'])
    translator = str.maketrans('', '', exclude)
    cleaned_text = cleaned_text.translate(translator)

    # Remove stopwords using SpaCy
    doc = nlp(cleaned_text)
    filtered_tokens = [token.text for token in doc if not token.is_stop]
    cleaned_text = ' '.join(filtered_tokens)

    return cleaned_text

def extract_names(text):
    # Rule Based Approach using Span
    # doc = nlp(text)
    # name_span = Span(doc, 0, 2, label="PERSON")
    # doc.set_ents([name_span], default="unmodified")
    # person_name = [ent.text for ent in doc.ents if ent.label_ == 'PERSON']

    # Entity Ruler Approach
    patterns = [
        [{'POS': 'PROPN'}, {'POS': 'PROPN'}],  # First name and Last name
        [{'POS': 'PROPN'}, {'POS': 'PROPN'}, {'POS': 'PROPN'}],  # First name, Middle name, and Last name
        [{'POS': 'PROPN'}, {'POS': 'PROPN'}, {'POS': 'PROPN'}, {'POS': 'PROPN'}]  # First name, Middle name, Middle name, and Last name
    ]

    matcher = Matcher(nlp.vocab)
    matcher.add("person_name", patterns)
    doc = nlp(text)

    # Extract entities matched by the Matcher
    matches = matcher(doc)
    person_name = [doc[start:end].text for match_id, start, end in matches]

    return person_name

def extract_contact_number(text):
    # Regex Pattern Approach
    # contact_numbers = []
    # pattern = r"\b(?:\+?\d{1,3}[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}\b"
    # match = re.search(pattern, text)
    # if match:
    #     contact_number = match.group()
    #     contact_numbers.append(contact_number)

    # Entity Ruler Approach
    contact_numbers = []
    ruler = nlp.add_pipe("entity_ruler")

    patterns = [
                    {"label": "PHONE_NUMBER", "pattern": [{"ORTH": "("}, {"SHAPE": "ddd"}, {"ORTH": ")"}, {"SHAPE": "ddd"},
                    {"ORTH": "-", "OP": "?"}, {"SHAPE": "dddd"}]}
                ]
    ruler.add_patterns(patterns)
    doc = nlp(text)

    # Extract entities matched by the entity_ruler
    for ent in doc.ents:
        if ent.label_ == "PHONE_NUMBER":
            contact_numbers.append(ent.text)

    return contact_numbers

def extract_email(text):
    email_address = []

    # Use regex pattern to find a potential email address
    pattern = r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b"
    match = re.search(pattern, text)
    if match:
        email = match.group()
        email_address.append(email)

    return email_address

def extract_skills(text):
    skills_list = ['Python', 'Java', 'C', 'C++', 'SQL', 'NoSQL', 'R', 'SAS', 'Tableau', 'Microsoft Excel',
                   'Power BI', 'Visio', 'MongoDB', 'Hadoop', 'Amazon Athena', 'Redshift', 'Glue',
                   'Alteryx', 'Apache Spark', 'QuickSight', 'Azure Synapse Studio', 'UNIX', 'Data Analysis',
                   'Federated Queries', 'Data Visualization', 'Business Intelligence', 'Cloud Computing',
                   'Data Modelling', 'Pattern and Trend Analysis', 'Shell Scripting', 'Hypothesis testing',
                   'Regression analysis', 'ANOVA', 'Time-series Forecasting', 'JIRA', 'MS Office',
                   'SharePoint', 'ServiceNow', 'MS SQL Server', 'Metabase', 'Confluence']
    skills = []

    for skill in skills_list:
        pattern = r"\b{}\b".format(re.escape(skill))
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            skills.append(skill)

    return skills 

def parse_resume(file_path, output_text_path): 
    try:
        # Extract text from the Word document
        text = docx2txt.process(file_path)

        # Text Pre-processing
        cleaned_text = text_preprocessing(text)

        # Extracting Person Name
        person_name = extract_names(cleaned_text)

        # Extracting Contact Information (Phone Number and Email Address)
        contact_numbers = extract_contact_number(cleaned_text)
        email_address = extract_email(cleaned_text)

        # Extracting Skills from Cleaned Text
        skills = extract_skills(cleaned_text)

        # Named Entity Recognition (NER) Modelling
        doc = nlp(cleaned_text)

        # Extract named entities
        named_entities = [(ent.text, ent.label_) for ent in doc.ents]

        # Write the cleaned text and named entities to the text file
        with open(output_text_path, 'w', encoding='utf-8') as txt_file:
            txt_file.write(cleaned_text + '\n\n')
            
            txt_file.write("Named Entities:\n")
            for entity, label in named_entities:
                txt_file.write(f"{entity} - {label}\n")

            txt_file.write("\nName: \n")
            for name in person_name:
                txt_file.write(f"{name}\n")

            txt_file.write("\nContact Information\n")
            txt_file.write("Phone Number: \n")
            for contact_number in contact_numbers:
                txt_file.write(f"{contact_number}\n")

            txt_file.write("\nEmail Address: \n")
            for email in email_address:
                txt_file.write(f"{email}\n")
            
            txt_file.write("\nSkills: \n")
            for skill in skills:
                txt_file.write(f"{skill}\n")

        print(f"Text extracted successfully with named entities and saved to: {output_text_path}")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    resume_file_path = '/Users/Ansh/Desktop/Python/Resume Parsing Project/Sample Resume.docx'
    output_text_path = '/Users/Ansh/Desktop/Python/Resume Parsing Project/output_text_file.txt'
    parse_resume(resume_file_path, output_text_path)
