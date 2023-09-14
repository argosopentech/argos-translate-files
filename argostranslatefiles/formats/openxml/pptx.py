from argostranslate.translate import ITranslation
from pptx import Presentation
import sys
from argostranslatefiles.formats.abstract_xml import AbstractXml


class Pptx(AbstractXml):
    supported_file_extensions = ['.pptx']

    def translate_paragraphs(self, paras, underlying_translation):
        for paragraph in paras:
            #get the average font size for the para
            if paragraph.runs:
                font_sizes = [(x.font.size,len(x.text)) for x in paragraph.runs]
                #sys.stderr.write(f"paragraph font size: {paragraph.font.size}. font sizes: {str(font_sizes)}\n")
                if font_sizes:
                    #avg_font_size = round(sum([x[0] for x in font_sizes])/len([x[0] for x in font_sizes]))
                    most_common_font_size = sorted(font_sizes,key=lambda x: x[1])[0]
                    if most_common_font_size:
                        paragraph.font.size = most_common_font_size[0]
                paragraph.text = underlying_translation.translate(paragraph.text)
         
    def translate(self, underlying_translation: ITranslation, file_path: str):
        outzip_path = self.get_output_path(underlying_translation, file_path)
        presentation = Presentation(file_path)
        
        for slide in presentation.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    self.translate_paragraphs(shape.text_frame.paragraphs, underlying_translation)
            
                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            self.translate_paragraphs(cell.text_frame.paragraphs, underlying_translation)
                            #sys.stderr.write(f"paragraph font size: {paragraph.font.size}. font sizes: {str(font_sizes)}\n")
                            #c.text = underlying_translation.translate(c.text)

        presentation.save(outzip_path)  


        return outzip_path
