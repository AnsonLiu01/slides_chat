import logging
from typing import List, Optional, Union
from pptx import Presentation
from sumy.parsers.plaintext import PlaintextParser
from sumy.nlp.tokenizers import Tokenizer
from sumy.summarizers.lsa import LsaSummarizer
import nltk

# Ensure necessary NLTK resources are downloaded
nltk.download('punkt')

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
log = logging.getLogger()

class SlidesIngest:
    """
    Class for ingesting and storing PowerPoint slides content.
    """
    def __init__(self, pp_file_loc: str):
        """
        :param pp_file_loc: PowerPoint file location
        """
        self.pp_file_loc = pp_file_loc
        self.prs = None
        self.slide_content = {}

    def load_file(self) -> None:
        """
        Function to load PowerPoint file and extract content.
        """
        log.info(f'Loading PowerPoint file: {self.pp_file_loc}')
        self.prs = Presentation(self.pp_file_loc)

        log.info(f'PowerPoint total slide count: {len(self.prs.slides)}')
        
        for slide_num, slide in enumerate(self.prs.slides, start=0):
            slide_text = []
            for shape in slide.shapes:
                if hasattr(shape, 'text') and shape.text.strip():
                    slide_text.append(shape.text)
            
            self.slide_content[slide_num] = " ".join(slide_text) if slide_text else ""

        log.info(f'Successfully extracted content from {len(self.slide_content)} slides.')
        
    def summarise(
        self, 
        n_slides: Optional[Union[int, List[int]]] = None
        ) -> None:
        """
        Function to summarise the PowerPoint slides
        :param n_slides: list of slides to summarise
        """
        log.info("Starting summarization process")
        
        n_slides = n_slides if n_slides else [n for n in range(len(self.prs.slides))]
        
        if isinstance(n_slides, int):
            n_slides = [n_slides]  # Convert a single slide number to a list
        elif isinstance(n_slides, list):
            if not all(isinstance(i, int) and 0 <= i < len(self.prs.slides) for i in n_slides):
                raise ValueError("Slide indices must be integers within the valid range")
        else:
            raise TypeError("n_slides must be an integer, a list of integers, or None")

        all_slide_text = "\n".join([f"Slide {n + 1}: {self.slide_content[n]}" for n in n_slides])

        parser = PlaintextParser.from_string(all_slide_text, Tokenizer("english"))
        summarizer = LsaSummarizer()
        
        summary = summarizer(parser.document, 8) 

        for sentence in summary:
            print(sentence)
        
    def runner(self) -> None:
        """
        Main runner function.
        """
        log.info(f'Beginning PowerPoint summarisation on file {self.pp_file_loc}')
        
        self.load_file()
        self.summarise()
        
        log.info('Summarisation Completed')

if __name__ == '__main__':
    slides = SlidesIngest(pp_file_loc='/Users/anson/Downloads/Lecture 10 Aggression 1 2020.pptx')
    slides.runner()

    for slide_num, content in slides.slide_content.items():
        print(f'SLIDE {slide_num}: {content}')
