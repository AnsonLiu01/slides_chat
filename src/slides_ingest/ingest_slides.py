from typing import List, Optional, Union, Tuple

import logging
from pptx import Presentation
from transformers import pipeline

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

log = logging.getLogger()


class SlidesIngest:
    """
    Class for ingesting and storing PowerPoint slides content.
    """
    def __init__(
        self, 
        pp_file_loc: str
        ):
        """
        :param pp_file_loc: PowerPoint file location
        """
        self.pp_file_loc = pp_file_loc
        
        self.short_sum = None
        self.long_sum = None
        
        self.prs = None
        self.slide_content = None
        self.slide_summary = None

    def init_summarisers(self) -> None:
        """
        Function to initialise all summariser tools
        """
        log.info('Initialisiing hugging face summary tools')
        
        self.short_sum = pipeline("summarization", model="sshleifer/distilbart-cnn-12-6")
        self.long_sum = pipeline("summarization", model="facebook/bart-large-cnn")
    
    @staticmethod
    def calc_min_max_tokens(
        input_length: int
    ) -> Tuple[int, int]:
        """
        Function to calculate the minimum and maximum tokens to use/cap at
        :param input_length: numbers of words in string
        :return: min and max token values
        """
        min_length = max(1, int(input_length * 0.1))
        max_length = max(5, min(input_length // 1.5, 200))

        return min_length, max_length

    def load_file(self) -> None:
        """
        Function to load PowerPoint file and extract content.
        """
        log.info(f'Loading PowerPoint file: {self.pp_file_loc}')
        
        self.slide_content = {}
        self.prs = Presentation(self.pp_file_loc)

        log.info(f'PowerPoint total slide count: {len(self.prs.slides)}')
        
        for slide_num, slide in enumerate(self.prs.slides, start=0):
            slide_text = []
            for shape in slide.shapes:
                if hasattr(shape, 'text') and shape.text.strip():
                    slide_text.append(shape.text)
            
            self.slide_content[slide_num] = " ".join(slide_text) if slide_text else ""

        log.info(f'Successfully extracted content from {len(self.slide_content)} slides.')

    def get_slide_text(
        self,
        n_slides: Optional[Union[int, List[int]]] = None
    ) -> str:
        """
        Function to get slides from all or a specific set of slides 
        :param n_slides: slide selection range, if None will get all
        :return all_slide_text: text from slide
        """
        log.info(f'Getting{" all " if n_slides is None else " "}slide text{f"" if n_slides is None else " in slides {n_slides}"}')
        
        all_slides = True if n_slides is None else False
        
        n_slides = n_slides if n_slides else [n for n in range(len(self.prs.slides))]
        n_slides = [n_slides] if isinstance(n_slides, int) else n_slides

        if all_slides: 
            all_slide_text = '\n'.join([self.slide_content[n] for n in n_slides])
        else:
            all_slide_text = '\n'.join([f"Slide {n + 1}: {self.slide_content[n]}" for n in n_slides])

        return all_slide_text

    def summarise_per_slide(
        self, 
        n_slides: Optional[Union[int, List[int]]] = None
    ) -> None:
        """
        Function to summarise each slide in deck.
        :param n_slides: list of slides to summarise; if None, summarises all slides
        """
        log.info("Summarising each slide")
        
        self.slide_summary = {}

        slides_to_summarize = n_slides if n_slides is not None else self.slide_content.keys()
        slides_to_summarize = [slides_to_summarize] if isinstance(slides_to_summarize, int) else slides_to_summarize
        
        for slide_no in slides_to_summarize:
            if slide_no in self.slide_content:
                slide_info = self.slide_content[slide_no]
                input_length = len(slide_info.split())
                
                if input_length != 0:
                    min_length, max_length = self.calc_min_max_tokens(input_length=input_length)
                    
                    pp_summary = self.short_sum(
                        slide_info, 
                        max_length=max_length, 
                        min_length=min_length, 
                        do_sample=False
                    )
                    
                    self.slide_summary[str(slide_no)] = pp_summary[0]['summary_text']

    @staticmethod
    def split_text_chunks(
        text: str, 
        max_token_length: int = 200
        ) -> List[str]:
        """
        Splits text into chunks that do not exceed the model's maximum token limit.
        :param text: Input text to split
        :param max_token_length: Token limit per chunk
        :return: List of text chunks
        """
        words = text.split()
        return [" ".join(words[i:i + max_token_length]) for i in range(0, len(words), max_token_length)]

    def summarise_all(self) -> None:
        """
        Function to summarise all slides as one, splitting into chunks if input exceeds model's token limit.
        """
        log.info('Summarising all slides')

        all_slides_text = self.get_slide_text()
        input_length = len(all_slides_text.split())

        # Check if text exceeds token limit
        if input_length > 1024:
            log.info("Splitting text into smaller chunks due to token length limit")
            text_chunks = self.split_text_chunks(all_slides_text)
            chunk_summaries = []
            
            n_chunk = 1

            for chunk in text_chunks:
                log.info(f'summarising chunk {n_chunk} of total {len(text_chunks)}')
                n_chunk += 1
                
                min_length, max_length = self.calc_min_max_tokens(input_length=len(chunk.split()))
                summary = self.short_sum(
                    chunk,
                    max_length=max_length,
                    min_length=min_length,
                    do_sample=False
                )[0]['summary_text']
                chunk_summaries.append(summary)
            
            combined_text = " ".join(chunk_summaries)
            
            self.slide_summary['all'] = combined_text
        else:
            # Summarize directly if within token limit
            min_length, max_length = self.calc_min_max_tokens(input_length=input_length)
            pp_summary = self.short_sum(
                all_slides_text,
                max_length=max_length,
                min_length=min_length,
                do_sample=False
            )
            self.slide_summary['all'] = pp_summary[0]['summary_text']
    
    def summarise_runner(self) -> None: 
        """
        Sub-runner function for all summarisation processes
        """
        # self.summarise_per_slide()
        self.summarise_all()
        
    def runner(self) -> None:
        """
        Main runner function.
        """
        log.info(f'Starting PowerPoint summarisation on file {self.pp_file_loc}')
        
        self.init_summarisers()
        self.load_file()
        
        self.summarise_runner()
        
        log.info('Summarisation Complete')

if __name__ == '__main__':
    slides = SlidesIngest(pp_file_loc='/Users/anson/Downloads/Lecture 10 Aggression 1 2020.pptx')
    slides.runner()
    