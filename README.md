PowerPoint Slide Ingestion and Summarization Tool (A WORK IN PROGRESS)

This repository contains a tool for ingesting PowerPoint slides and generating concise summaries. It leverages NLP techniques to extract key insights from slide content, making it ideal for quick reviews and analysis.

Features

Slide Ingestion: Reads PowerPoint slides and processes text content.
Text Summarization: Uses NLP models to generate summaries of slide content.
Flexible Output Options: Allows display in the console or saving to a file.
Getting Started

Prerequisites
Python 3.9 or later
Required Python packages (install via requirements.txt)
Installation
Clone the repository:
git clone https://github.com/AnsonLiu01/slides_chat.git
cd powerpoint-ingestion-tool
Install dependencies:
pip install -r requirements.txt
Usage
Place the PowerPoint file you want to process in the project directory (e.g., Week 2 Friday.pptx).
Run the main script with the following command:
python main.py
Adjust options in main.py:
save: Set to True if you want to save the summary to a file.
display: Set to True if you want the summary to display in the console.
Code Overview
main.py: This is the entry point for running the tool. It imports the SlidesIngest class and initiates the ingestion and summarization process.
slides_ingest/ingest_slides.py: Contains the SlidesIngest class, which handles reading and processing slide content.
Example
Here’s an example of how to use the tool in main.py:
  Change the pp_filename to the filename of the powerpoint with the suffix 'pptx' 

    slides = SlidesIngest(pp_filename='Week 2 Friday.pptx')
    
    slides.runner(
        save=False, 
        display=True
    )
  
Output
If display=True, the summary will appear in the console. If save=True, it will save to a specified file path.

Contributors
Anson
