import pdfplumber
import re
import string
import numpy as np
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import nltk
from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer
import os
import docx

# Download necessary NLTK data files
nltk.download('stopwords')
nltk.download('wordnet')

# Function to preprocess text data with advanced techniques
def advanced_preprocess_text(text):
    # Remove unwanted characters and convert to lowercase
    text = re.sub(r'[^a-zA-Z0-9\s]', '', text.lower())
    
    # Remove extra whitespaces
    text = ' '.join(text.split())
    
    # Tokenize and remove stop words
    stop_words = set(stopwords.words('english'))
    tokens = text.split()
    tokens = [word for word in tokens if word not in stop_words]
    
    # Lemmatize tokens
    lemmatizer = WordNetLemmatizer()
    tokens = [lemmatizer.lemmatize(word) for word in tokens]
    
    return ' '.join(tokens)

# Function to extract text from PDF files
def extract_text_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = ''
        for page in pdf.pages:
            text += page.extract_text()
    return text

# Function to extract text from Word documents
def extract_text_from_word(docx_path):
    doc = docx.Document(docx_path)
    text = ''
    for paragraph in doc.paragraphs:
        text += paragraph.text
    return text

# Determine the file type and extract text accordingly
def extract_text(file_path):
    file_extension = os.path.splitext(file_path)[1].lower()
    if file_extension == '.pdf':
        return extract_text_from_pdf(file_path)
    elif file_extension == '.docx':
        return extract_text_from_word(file_path)
    else:
        raise ValueError(f"Unsupported file type: {file_extension}")

# File paths for job description and CVs
job_description_path = r"C:\Users\carso\Documents\July\new cvs\digital marketers.pdf"
cv_paths = [
    r"C:\Users\carso\Documents\July\new cvs\batch 2 digital marketers\EvalyneWanjiku-CV.pdf",
r"C:\Users\carso\Documents\July\new cvs\batch 2 digital marketers\FaithGitonga-CV.pdf",
r"C:\Users\carso\Documents\July\new cvs\batch 2 digital marketers\khasoavivian-CV.pdf",
r"C:\Users\carso\Documents\July\new cvs\batch 2 digital marketers\NicholasMuraya-CV.pdf",
r"C:\Users\carso\Documents\July\new cvs\batch 2 digital marketers\NicoleKigame-CV.docx",
r"C:\Users\carso\Documents\July\new cvs\batch 2 digital marketers\PurityNgigi-CV (1).docx",
r"C:\Users\carso\Documents\July\new cvs\batch 2 digital marketers\WINNIESABINAOMUYA-CV.docx",
        # ... add more CV file paths here
]

# Extract text from job description
job_description_text = extract_text(job_description_path)

# Extract text from CVs
cv_texts = [extract_text(cv_path) for cv_path in cv_paths]

# Preprocess the job description and CVs using advanced preprocessing
job_description = advanced_preprocess_text(job_description_text)
cvs = [advanced_preprocess_text(cv_text) for cv_text in cv_texts]

# Create TF-IDF vectors
vectorizer = TfidfVectorizer()
job_description_vector = vectorizer.fit_transform([job_description])
cv_vectors = vectorizer.transform(cvs)

# Calculate cosine similarity between job description and CVs
similarities = cosine_similarity(job_description_vector, cv_vectors).flatten()

# Extract candidate names from file paths
candidate_names = [os.path.splitext(os.path.basename(cv_path))[0] for cv_path in cv_paths]

# Create a dataframe with candidate names and their similarity scores
cv_scores = pd.DataFrame({'CV': candidate_names, 'Similarity': similarities})

# Sort the dataframe based on similarity scores (descending order)
cv_scores = cv_scores.sort_values(by='Similarity', ascending=False)

# Save the dataframe as an Excel file
cv_scores.to_excel('Digital Marketers batch 2.xlsx', index=False)

# Print the top 5 candidates
print("Top 10 Candidates:")
print(cv_scores.head(5))
