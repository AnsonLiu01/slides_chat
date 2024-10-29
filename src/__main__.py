from src.slides_ingest.ingest_slides import SlidesIngest

if __name__ == '__main__':
    
    slides = SlidesIngest(pp_file_loc='')
    
    slides.runner()