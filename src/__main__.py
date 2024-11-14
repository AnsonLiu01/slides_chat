import os

from slides_ingest.ingest_slides import SlidesIngest

if __name__ == '__main__':
    os.environ["TOKENIZERS_PARALLELISM"] = 'false'

    slides = SlidesIngest(pp_filename='Week 2 Friday.pptx')
    
    slides.runner(
        save=False, 
        display=True
        )
    