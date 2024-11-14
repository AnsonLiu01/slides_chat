import os

from slides_ingest.ingest_slides import SlidesIngest

if __name__ == '__main__':
    os.environ["TOKENIZERS_PARALLELISM"] = 'false'

    slides = SlidesIngest(pp_filename='8a815e94056b8572942ec6bfb545d78bweek_1_friday_400hsc_block_2.pptx')
    
    slides.runner(
        save=False, 
        display=True
        )
    