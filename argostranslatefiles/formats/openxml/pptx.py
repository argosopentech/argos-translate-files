from argostranslate.translate import ITranslation
from pptx import Presentation

from argostranslatefiles.formats.abstract_xml import AbstractXml


class Pptx(AbstractXml):
    supported_file_extensions = ['.pptx']

    def translate(self, underlying_translation: ITranslation, file_path: str):
        outzip_path = self.get_output_path(underlying_translation, file_path)
        presentation = Presentation(file_path)
        
        for slide in presentation.slides:
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                for paragraph in shape.text_frame.paragraphs:
                    paragraph.text = underlying_translation.translate(paragraph.text)
       
        presentation.save(outzip_path)  


        return outzip_path
