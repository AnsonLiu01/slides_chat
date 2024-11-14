from slides_ingest.ingest_slides import SlidesIngest

if __name__ == '__main__':
    slides = SlidesIngest(pp_filename='Precision Medicine Lecture 1_BB_slides only.pptx')
    
    slides.runner(
        save=False, 
        display=True
        )
    