import pandas as pd
import re

# ------------------------------------------------------------------
# 1. Load the Excel file
# ------------------------------------------------------------------
input_file = "elsevier_citation.xlsx"
df = pd.read_excel(input_file, sheet_name="Sheet1")

# ------------------------------------------------------------------
# 2. Define keyword sets (case‑insensitive)
# ------------------------------------------------------------------
ai_terms = [
    "artificial intelligence", "machine learning", "deep learning",
    "neural network", "ai-driven", "intelligent systems",
    "support vector machine", "random forest", "gradient boosting",
    "xgboost", "lstm", "gru", "autoencoder", "generative adversarial",
    "isolation forest", "one-class", "ensemble learning", "transfer learning",
    "federated learning", "reinforcement learning", "bayesian network",
    "supervised learning", "unsupervised learning", "semi-supervised"
]

fraud_terms = [
    "fraud detection", "fraud prevention", "anomaly detection",
    "fraudulent transaction", "financial fraud", "credit card fraud",
    "payment fraud", "transaction fraud", "fraudulent activity",
    "anti-fraud", "fraud risk", "money laundering", "default detection"
]

card_terms = [
    "smart card", "chip card", "emv", "contactless payment",
    "payment card", "debit card", "credit card", "pos transaction",
    "credit card fraud", "card transaction", "card payment",
    "prepaid card", "stored-value card", "bank card", "plastic card"
]

# Pre‑compile regex (word boundaries)
ai_re = re.compile(r'\b(?:' + '|'.join(re.escape(term) for term in ai_terms) + r')\b', re.IGNORECASE)
fraud_re = re.compile(r'\b(?:' + '|'.join(re.escape(term) for term in fraud_terms) + r')\b', re.IGNORECASE)
card_re = re.compile(r'\b(?:' + '|'.join(re.escape(term) for term in card_terms) + r')\b', re.IGNORECASE)

# ------------------------------------------------------------------
# 3. Function to classify each row
# ------------------------------------------------------------------
def classify_row(row):
    # Combine citation (column B) and abstract (column D)
    text = f"{row['citation']} {row['abstract']}".lower()

    has_ai = bool(ai_re.search(text))
    has_fraud = bool(fraud_re.search(text))
    has_card = bool(card_re.search(text))

    if has_ai and has_fraud and has_card:
        return ("Include", "Contains AI/ML, fraud detection, and payment card terms.")
    elif has_ai and has_fraud and not has_card:
        return ("Borderline", "Contains AI/ML and fraud terms, but no explicit payment card terminology.")
    elif has_ai and has_card and not has_fraud:
        return ("Borderline", "Contains AI/ML and payment card terms, but no explicit fraud detection focus.")
    elif has_fraud and has_card and not has_ai:
        return ("Borderline", "Contains fraud and payment card terms, but no AI/ML methods.")
    else:
        return ("Exclude", "Does not simultaneously address AI/ML, fraud detection, and payment cards.")

# Apply classification
df[['Decision', 'Reason']] = df.apply(classify_row, axis=1, result_type='expand')

# ------------------------------------------------------------------
# 4. Save the annotated file
# ------------------------------------------------------------------
output_file = "elsevier_citation_annotated.xlsx"
df.to_excel(output_file, index=False)
print(f"Annotated file saved to: {output_file}")
print("\nDecision counts:")
print(df['Decision'].value_counts())
