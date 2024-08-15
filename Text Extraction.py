import docx2txt
import re
import spacy

# Path to the downloaded language model folder
spacy_model_path = '/Users/Ansh/Downloads/en_core_web_sm-3.7.1/en_core_web_sm/en_core_web_sm-3.7.1'
nlp = spacy.load(spacy_model_path)

def parse_resume(file_path, output_text_path): 
    try:
        # Extract text from the Word document
        text = docx2txt.process(file_path)

        # Removing extra spaces and empty lines
        cleaned_text = re.sub(r'\s+', ' ', text).strip()

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

        print(f"Text extracted successfully with named entities and saved to: {output_text_path}")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    resume_file_path = '/Users/Ansh/Desktop/Python/Resume Parsing Project/Sample Resumes/Sample_Resume.docx'
    output_text_path = '/Users/Ansh/Desktop/Python/Resume Parsing Project/Output Text Files (NER)/Sample_Resume.txt'
    parse_resume(resume_file_path, output_text_path)
