import pdfplumber
import re
import string
import numpy as np
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# Function to preprocess text data
def preprocess_text(text):
    # Remove unwanted characters and convert to lowercase
    text = re.sub(r'[^a-zA-Z0-9\s]', '', text.lower())
    
    # Remove extra whitespaces
    text = ' '.join(text.split())
    
    return text

# Function to extract text from PDF files
def extract_text_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = ''
        for page in pdf.pages:
            text += page.extract_text()
    return text

# File paths for job description and CVs
job_description_path = r"C:\Users\carso\Documents\June\Week 2\Business Developer JD.pdf"
cv_paths = [
    
    r"C:\Users\carso\Documents\June\Batch 2\BeatriceNyagilo-CV.pdf",
r"C:\Users\carso\Documents\June\Batch 2\CollinsMukhebi-CV.pdf",
r"C:\Users\carso\Documents\June\Batch 2\EmilyWawira-CV.pdf",
r"C:\Users\carso\Documents\June\Batch 2\EMMANUELKINYUAMUKEMBU-CV.pdf",
r"C:\Users\carso\Documents\June\Batch 2\JaneMaina-CV.pdf",
r"C:\Users\carso\Documents\June\Batch 2\JosephMutinda-CV.pdf",
r"C:\Users\carso\Documents\June\Batch 2\LynnetteWaweru-CV.pdf",
r"C:\Users\carso\Documents\June\Batch 2\MarvinNjeru-CV.pdf",
r"C:\Users\carso\Documents\June\Batch 2\marychege-CV.pdf",
r"C:\Users\carso\Documents\June\Batch 2\meshacklumumba-CV.pdf",
r"C:\Users\carso\Documents\June\Batch 2\MichaelNyokabi-CV.pdf",
r"C:\Users\carso\Documents\June\Batch 2\moseskigen-CV.pdf",
r"C:\Users\carso\Documents\June\Batch 2\MosesNjiru-CV.pdf",
r"C:\Users\carso\Documents\June\Batch 2\PHILIPAGANDI-CV.pdf",
r"C:\Users\carso\Documents\June\Batch 2\ReginaGikera-CV.pdf",
r"C:\Users\carso\Documents\June\Batch 2\StephenMayaka-CV.pdf"

    # ... add more CV file paths here
]

# Extract text from job description PDF
job_description_text = extract_text_from_pdf(job_description_path)

# Extract text from CV PDFs
cv_texts = [extract_text_from_pdf(cv_path) for cv_path in cv_paths]

# Preprocess the job description and CVs
job_description = preprocess_text(job_description_text)
cvs = [preprocess_text(cv_text) for cv_text in cv_texts]

# Create TF-IDF vectors
vectorizer = TfidfVectorizer()
job_description_vector = vectorizer.fit_transform([job_description])
cv_vectors = vectorizer.transform(cvs)

# Calculate cosine similarity between job description and CVs
similarities = cosine_similarity(job_description_vector, cv_vectors).flatten()

# Create a dataframe with CV indices and their similarity scores
cv_scores = pd.DataFrame({'CV': [f'Candidate {i+1}' for i in range(len(cv_paths))], 'Similarity': similarities})

# Sort the dataframe based on similarity scores (descending order)
cv_scores = cv_scores.sort_values(by='Similarity', ascending=False)

# Save the dataframe as an Excel file
cv_scores.to_excel('cv_similarity_scores2.xlsx', index=False)
