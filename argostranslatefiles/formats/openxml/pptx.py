from argostranslate.translate import ITranslation
from pptx import Presentation
import sys
from argostranslatefiles.formats.abstract_xml import AbstractXml


class Pptx(AbstractXml):
    supported_file_extensions = ['.pptx']

    def translate_paragraphs(self, paras, underlying_translation):
        for paragraph in paras:
            paragraph.text = underlying_translation.translate(paragraph.text)
        
    def translate(self, underlying_translation: ITranslation, file_path: str):
        outzip_path = self.get_output_path(underlying_translation, file_path)
        presentation = Presentation(file_path)
        
        for slide in presentation.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    self.translate_paragraphs(shape.text_frame.paragraphs, underlying_translation)
            
                if shape.has_table:
                    for r in shape.table.rows:
                        for c in r.cells:
                            c.text = underlying_translation.translate(c.text)

        presentation.save(outzip_path)  


        return outzip_path
