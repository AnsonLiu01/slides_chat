from transformers import pipeline

# Load a pre-trained summarization model
summarizer = pipeline("summarization", model="sshleifer/distilbart-cnn-12-6")

# Input text to summarize
text = 'Your long text here that needs summarizing...'
    
# Perform summarization
summary = summarizer(text, max_length=130, min_length=30, do_sample=False)

# Output the summary
print(f'ANSON THE SUMMARISATION IS HERE: {summary[0]["summary_text"]}')
