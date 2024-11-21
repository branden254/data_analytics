import pdfplumber
import re
import string
import numpy as np
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import os

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
job_description_path = r'C:\Users\carso\Documents\June\Business Developer JD.pdf'
cv_paths = [
    r'C:\Users\carso\Documents\June\Batch 2\AbrahamNgetich-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\AlexanderMumo-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\AlfayoKemei-CV.pdf',
    r"C:\Users\carso\Documents\June\Batch 2\AmosMuthui-CV.pdf",
    r"C:\Users\carso\Documents\June\Batch 2\anitahShivachi-CV.pdf",
    r"C:\Users\carso\Documents\June\Batch 2\AnneKibiro-CV.pdf",
    r"C:\Users\carso\Documents\June\Batch 2\AnthonyLangat-CV.pdf",
    r"C:\Users\carso\Documents\June\Batch 2\BonifaceNjenga-CV.pdf",
    r"C:\Users\carso\Documents\June\Batch 2\BRENDAOTIENO-CV.pdf",
    r"C:\Users\carso\Documents\June\Batch 2\CALEBSAMBRIR-CV.pdf",
    r"C:\Users\carso\Documents\June\Batch 2\CarolineMaina-CV.pdf",
    r"C:\Users\carso\Documents\June\Batch 2\catherinengendo-CV.pdf",
    r"C:\Users\carso\Documents\June\Batch 2\CatherineWambua-CV.pdf",
    r"C:\Users\carso\Documents\June\Batch 2\charlesmabeya-CV.pdf",
    r"C:\Users\carso\Documents\June\Batch 2\DanielNdegwa-CV.pdf",
    r"C:\Users\carso\Documents\June\Batch 2\DELBERTWanjala-CV.pdf",
    r"C:\Users\carso\Documents\June\Batch 2\DennisAmukhuma-CV.pdf",
    r"C:\Users\carso\Documents\June\Batch 2\DennisKariuki-CV.pdf",
    r"C:\Users\carso\Documents\June\Batch 2\DerrickPhillipwafula-CV.pdf",
    r"C:\Users\carso\Documents\June\Batch 2\DismasObat-CV.pdf",
    r"C:\Users\carso\Documents\June\Batch 2\DorcasChepkirui-CV.pdf",
    r"C:\Users\carso\Documents\June\Batch 2\DorriceAkinyi-CV.pdf",
    r"C:\Users\carso\Documents\June\Batch 2\EdwinochiengOgundo-CV.pdf",
    r"C:\Users\carso\Documents\June\Batch 2\EdwinOsore-CV.pdf",
    r"C:\Users\carso\Documents\June\Batch 2\EliasSande-CV.pdf",
    r"C:\Users\carso\Documents\June\Batch 2\EmmanuelMosonik-CV.pdf",
    r"C:\Users\carso\Documents\June\Batch 2\EnockOmondi-CV.pdf",
    r'C:\Users\carso\Documents\June\Batch 2\ephaholendo-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\FaithAmondi-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\FaithOyoo-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\FranklineKaraniA-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\GeoffreyKonya-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\GeorgeKiptoo-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\GeorgeMutiso-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\GeorgeOmollo-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\GetrudeMukami-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\GithukiJoseph-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\godfreyPalia-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\GraceGunya-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\GRACENYAMBURA-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\HabibWanjala-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\HerinaWanjala-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\HoseaNjoroge-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\IanIhaji-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\INNOCENTKHABI-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\JamesGachanja-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\JamesMBURU-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\JamesMutunga-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\JAMESMWIRIGIjames-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\JamesNjoroge-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\JamesOwino-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\JanetOtuoro-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\JohnMugo-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\JosephMunjoguWainaina-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\KapeenJackson-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\kennedybarasa-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\kennethkipchumba-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\kevinkimani-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\KevinOdhiambo-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\KibetDennis-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\LillianWambura-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\luckyngoya-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\madagadickson-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\MarkKaranja-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\MarkpaulNdirangu-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\MartinIreri-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\martinkamau-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\MaryMbuthia-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\MichaelKithambo-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\mildredkoech-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\MILKAHONESMUS-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\MoureenOwino-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\MuchaiKarera-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\MumoKivindu-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\NicodemusMageka-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\NORBERTMUGENI-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\OdayaGordon-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\OmondiSymon-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\PascalineKanini-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\pascalmusyoki-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\patrickkipchumba-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\PATRICKMUTOONI-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\PaulAbidha-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\PETEROSANO-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\PhilipWaruingi-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\PHINEASTHURANIRA-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\RobbinMasuti-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\RobinsonAminga-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\RodneyGuga-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\RoseVamba-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\RotichBenjamin-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\SamuelMusyoki-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\SamuelOtwori-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\SamuelRaongo-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\sharonJoseph-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\SimonNdirangu-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\SusanMwangi-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\SydneyOchieng-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\TeresiahNyambura-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\ThomasAmondo-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\TruphenahLiyala-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\VeronicahAmollo-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\VictorTitus-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\vincentambeva-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\VincentOmwayo-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\VinnyOtach-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\vivianotieno-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\VollineOkaka-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\WesaDuncan-CV.pdf',
    r'C:\Users\carso\Documents\June\Batch 2\WilsonMartha-CV.pdf'



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

# Extract names from file paths
cv_names = [os.path.basename(cv_path).split('-CV.pdf')[0] for cv_path in cv_paths]

# Create a dataframe with CV names and their similarity scores
cv_scores = pd.DataFrame({'CV': cv_names, 'Similarity': similarities})

# Sort the dataframe based on similarity scores (descending order)
cv_scores = cv_scores.sort_values(by='Similarity', ascending=False)

# Save the dataframe as an Excel file
cv_scores.to_excel('cv_similarity_scores3.xlsx', index=False)
